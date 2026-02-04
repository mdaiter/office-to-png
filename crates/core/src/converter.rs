//! Main converter orchestrator that ties together the LibreOffice pool and PDF renderer.
//!
//! This module provides the high-level API for converting Office documents to PNG images,
//! with support for batch processing, progress callbacks, and streaming output.

use crate::config::{
    BatchResult, ConversionProgress, ConversionRequest, ConversionStage, ConverterConfig,
    FailedFile, FileResult, PngPage, RenderConfig,
};
use crate::error::{ConversionError, Result};
use crate::pdf_renderer::PdfRenderer;
use crate::pool::LibreOfficePool;
use futures::stream::{self, Stream, StreamExt};
use std::path::{Path, PathBuf};
use std::sync::Arc;
use std::time::{Duration, Instant};
use tokio::sync::mpsc;
use tracing::{debug, error, info, warn};

/// Main converter for Office documents to PNG.
///
/// This is the primary interface for converting documents. It manages a pool of
/// LibreOffice instances and a PDF renderer for optimal throughput.
pub struct Converter {
    /// LibreOffice process pool.
    pool: Arc<LibreOfficePool>,
    /// PDF to PNG renderer.
    renderer: Arc<PdfRenderer>,
    /// Configuration.
    config: ConverterConfig,
}

impl Converter {
    /// Create a new converter with the given configuration.
    pub async fn new(config: ConverterConfig) -> Result<Self> {
        config.validate()?;

        info!(
            "Initializing converter with pool_size={}, dpi={}",
            config.pool.pool_size, config.render.dpi
        );

        let pool = LibreOfficePool::new(config.pool.clone()).await?;
        let renderer = PdfRenderer::new(config.render.clone())?;

        Ok(Self {
            pool: Arc::new(pool),
            renderer: Arc::new(renderer),
            config,
        })
    }

    /// Create a converter with default settings.
    pub async fn default() -> Result<Self> {
        Self::new(ConverterConfig::default()).await
    }

    /// Convert a single document to PNG images.
    pub async fn convert(&self, request: ConversionRequest) -> Result<FileResult> {
        let start = Instant::now();
        let input_path = request.input_path.clone();

        info!("Converting {:?}", input_path);

        // Stage 1: Convert to PDF via LibreOffice
        debug!("Stage 1: Converting to PDF");
        let pdf_path = self.pool.convert_to_pdf(&request.input_path).await?;

        // Stage 2: Render PDF pages to PNG
        debug!("Stage 2: Rendering PDF to PNG");
        let dpi = request.dpi_override.unwrap_or(self.config.render.dpi);
        let prefix = request.get_output_prefix();

        // Use the existing renderer with the specified DPI
        let pages = self.renderer
            .render_and_save_with_dpi(&pdf_path, &request.output_dir, &prefix, dpi)?;

        // Clean up temp PDF
        if let Err(e) = std::fs::remove_file(&pdf_path) {
            warn!("Failed to remove temp PDF {:?}: {}", pdf_path, e);
        }

        let output_paths: Vec<PathBuf> = pages
            .iter()
            .filter_map(|p| p.output_path.clone())
            .collect();

        info!(
            "Converted {:?} to {} pages in {:?}",
            input_path,
            pages.len(),
            start.elapsed()
        );

        Ok(FileResult {
            input_path,
            output_paths,
            page_count: pages.len(),
            duration: start.elapsed(),
        })
    }

    /// Convert multiple documents in batch.
    pub async fn convert_batch(&self, requests: Vec<ConversionRequest>) -> BatchResult {
        let start = Instant::now();
        let total_files = requests.len();
        let mut successful = Vec::new();
        let mut failed = Vec::new();
        let mut total_pages = 0;

        for request in requests {
            let input_path = request.input_path.clone();
            match self.convert(request).await {
                Ok(result) => {
                    total_pages += result.page_count;
                    successful.push(result);
                }
                Err(e) => {
                    error!("Failed to convert {:?}: {}", input_path, e);
                    failed.push(FailedFile {
                        input_path,
                        error: e.to_string(),
                    });
                }
            }
        }

        BatchResult {
            successful,
            failed,
            total_duration: start.elapsed(),
            total_pages,
        }
    }

    /// Convert multiple documents with progress callback.
    pub async fn convert_batch_with_progress<F>(
        &self,
        requests: Vec<ConversionRequest>,
        progress_callback: F,
    ) -> BatchResult
    where
        F: Fn(ConversionProgress) + Send + Sync,
    {
        let start = Instant::now();
        let total_files = requests.len();
        let mut successful = Vec::new();
        let mut failed = Vec::new();
        let mut total_pages = 0;

        for (file_index, request) in requests.into_iter().enumerate() {
            let input_path = request.input_path.clone();
            let current_file = input_path
                .file_name()
                .and_then(|n| n.to_str())
                .unwrap_or("unknown")
                .to_string();

            // Report starting
            progress_callback(ConversionProgress {
                file_index,
                total_files,
                current_file: current_file.clone(),
                pages_completed: 0,
                total_pages: None,
                stage: ConversionStage::ConvertingToPdf,
            });

            match self.convert_with_page_progress(&request, |pages_done, total| {
                progress_callback(ConversionProgress {
                    file_index,
                    total_files,
                    current_file: current_file.clone(),
                    pages_completed: pages_done,
                    total_pages: Some(total),
                    stage: ConversionStage::RenderingPages,
                });
            }).await {
                Ok(result) => {
                    total_pages += result.page_count;

                    // Report completion
                    progress_callback(ConversionProgress {
                        file_index,
                        total_files,
                        current_file: current_file.clone(),
                        pages_completed: result.page_count,
                        total_pages: Some(result.page_count),
                        stage: ConversionStage::Completed,
                    });

                    successful.push(result);
                }
                Err(e) => {
                    error!("Failed to convert {:?}: {}", input_path, e);

                    // Report failure
                    progress_callback(ConversionProgress {
                        file_index,
                        total_files,
                        current_file: current_file.clone(),
                        pages_completed: 0,
                        total_pages: None,
                        stage: ConversionStage::Failed,
                    });

                    failed.push(FailedFile {
                        input_path,
                        error: e.to_string(),
                    });
                }
            }
        }

        BatchResult {
            successful,
            failed,
            total_duration: start.elapsed(),
            total_pages,
        }
    }

    /// Convert with per-page progress reporting.
    async fn convert_with_page_progress<F>(
        &self,
        request: &ConversionRequest,
        page_callback: F,
    ) -> Result<FileResult>
    where
        F: Fn(usize, usize),
    {
        let start = Instant::now();
        let input_path = request.input_path.clone();

        // Stage 1: Convert to PDF
        let pdf_path = self.pool.convert_to_pdf(&request.input_path).await?;

        // Stage 2: Render with progress
        let dpi = request.dpi_override.unwrap_or(self.config.render.dpi);
        let prefix = request.get_output_prefix();

        // Ensure output directory exists
        std::fs::create_dir_all(&request.output_dir).map_err(|e| {
            ConversionError::OutputDirError {
                path: request.output_dir.clone(),
                message: e.to_string(),
            }
        })?;

        // Use streaming renderer for progress
        let renderer = if dpi != self.config.render.dpi {
            Arc::new(PdfRenderer::new(RenderConfig::with_dpi(dpi))?)
        } else {
            Arc::clone(&self.renderer)
        };

        let iter = renderer.render_pages_iter(&pdf_path)?;
        let total_pages = iter.len();
        let mut pages = Vec::with_capacity(total_pages);
        let mut pages_completed = 0;

        for page_result in iter {
            let mut page = page_result?;

            // Save to disk
            let filename = format!("{}_page_{:04}.png", prefix, page.page_number);
            let output_path = request.output_dir.join(&filename);
            std::fs::write(&output_path, &page.data).map_err(|e| {
                ConversionError::OutputDirError {
                    path: output_path.clone(),
                    message: e.to_string(),
                }
            })?;
            page.output_path = Some(output_path);

            pages_completed += 1;
            page_callback(pages_completed, total_pages);

            pages.push(page);
        }

        // Clean up temp PDF
        if let Err(e) = std::fs::remove_file(&pdf_path) {
            warn!("Failed to remove temp PDF {:?}: {}", pdf_path, e);
        }

        let output_paths: Vec<PathBuf> = pages
            .iter()
            .filter_map(|p| p.output_path.clone())
            .collect();

        Ok(FileResult {
            input_path,
            output_paths,
            page_count: pages.len(),
            duration: start.elapsed(),
        })
    }

    /// Convert a document and stream pages as they're rendered.
    ///
    /// This is useful for processing large documents where you want to start
    /// handling output before the entire document is processed.
    pub fn convert_stream(
        &self,
        request: ConversionRequest,
    ) -> impl Stream<Item = Result<PngPage>> + '_ {
        let pool = Arc::clone(&self.pool);
        let renderer = Arc::clone(&self.renderer);
        let config = self.config.clone();

        stream::once(async move {
            // First, convert to PDF
            let pdf_path = pool.convert_to_pdf(&request.input_path).await?;

            // Then create renderer with appropriate DPI
            let dpi = request.dpi_override.unwrap_or(config.render.dpi);
            let render_config = RenderConfig::with_dpi(dpi);
            let temp_renderer = PdfRenderer::new(render_config)?;

            // Return the PDF path and renderer for streaming
            Ok::<_, ConversionError>((pdf_path, temp_renderer, request))
        })
        .flat_map(|result| {
            match result {
                Ok((pdf_path, renderer, request)) => {
                    // Create a stream from the page iterator
                    let prefix = request.get_output_prefix();
                    let output_dir = request.output_dir.clone();

                    match renderer.render_pages_iter(&pdf_path) {
                        Ok(iter) => {
                            let pages: Vec<_> = iter
                                .enumerate()
                                .map(move |(idx, page_result)| {
                                    page_result.and_then(|mut page| {
                                        // Optionally save to disk
                                        let filename = format!(
                                            "{}_page_{:04}.png",
                                            prefix, page.page_number
                                        );
                                        let output_path = output_dir.join(&filename);
                                        std::fs::write(&output_path, &page.data).map_err(|e| {
                                            ConversionError::OutputDirError {
                                                path: output_path.clone(),
                                                message: e.to_string(),
                                            }
                                        })?;
                                        page.output_path = Some(output_path);
                                        Ok(page)
                                    })
                                })
                                .collect();

                            // Clean up temp PDF (ignore errors)
                            let _ = std::fs::remove_file(&pdf_path);

                            stream::iter(pages).boxed()
                        }
                        Err(e) => stream::once(async move { Err(e) }).boxed(),
                    }
                }
                Err(e) => stream::once(async move { Err(e) }).boxed(),
            }
        })
    }

    /// Convert documents in parallel batches.
    ///
    /// This processes `concurrency` documents simultaneously for maximum throughput.
    pub async fn convert_parallel(
        &self,
        requests: Vec<ConversionRequest>,
        concurrency: usize,
    ) -> BatchResult {
        let start = Instant::now();
        let total_files = requests.len();

        let results: Vec<(ConversionRequest, Result<FileResult>)> = stream::iter(requests)
            .map(|request| {
                let input_path = request.input_path.clone();
                async move {
                    let result = self.convert(request.clone()).await;
                    (request, result)
                }
            })
            .buffer_unordered(concurrency)
            .collect()
            .await;

        let mut successful = Vec::new();
        let mut failed = Vec::new();
        let mut total_pages = 0;

        for (request, result) in results {
            match result {
                Ok(file_result) => {
                    total_pages += file_result.page_count;
                    successful.push(file_result);
                }
                Err(e) => {
                    failed.push(FailedFile {
                        input_path: request.input_path,
                        error: e.to_string(),
                    });
                }
            }
        }

        BatchResult {
            successful,
            failed,
            total_duration: start.elapsed(),
            total_pages,
        }
    }

    /// Get pool health information.
    pub async fn health(&self) -> crate::pool::PoolHealth {
        self.pool.health().await
    }

    /// Shutdown the converter and release resources.
    pub async fn shutdown(&self) {
        info!("Shutting down converter");
        self.pool.shutdown().await;
    }

    /// Get the current configuration.
    pub fn config(&self) -> &ConverterConfig {
        &self.config
    }

    /// Get statistics about processing.
    pub fn stats(&self) -> ConverterStats {
        ConverterStats {
            total_documents_processed: self.pool.total_processed(),
            pool_size: self.config.pool.pool_size,
            dpi: self.config.render.dpi,
        }
    }
}

/// Statistics about the converter.
#[derive(Debug, Clone)]
pub struct ConverterStats {
    /// Total documents processed since creation.
    pub total_documents_processed: usize,
    /// Pool size.
    pub pool_size: usize,
    /// Configured DPI.
    pub dpi: u32,
}

/// Builder for creating a Converter with custom settings.
pub struct ConverterBuilder {
    config: ConverterConfig,
}

impl ConverterBuilder {
    /// Create a new builder with default settings.
    pub fn new() -> Self {
        Self {
            config: ConverterConfig::default(),
        }
    }

    /// Set the pool size.
    pub fn pool_size(mut self, size: usize) -> Self {
        self.config.pool.pool_size = size;
        self
    }

    /// Set the DPI for rendering.
    pub fn dpi(mut self, dpi: u32) -> Self {
        self.config.render.dpi = dpi;
        self
    }

    /// Set the conversion timeout.
    pub fn conversion_timeout(mut self, timeout: Duration) -> Self {
        self.config.pool.conversion_timeout = timeout;
        self
    }

    /// Set the number of render threads.
    pub fn render_threads(mut self, threads: usize) -> Self {
        self.config.render.render_threads = threads;
        self
    }

    /// Set the path to soffice binary.
    pub fn soffice_path(mut self, path: PathBuf) -> Self {
        self.config.pool.soffice_path = Some(path);
        self
    }

    /// Set the temporary directory.
    pub fn temp_dir(mut self, dir: PathBuf) -> Self {
        self.config.pool.temp_dir = Some(dir);
        self
    }

    /// Build the converter.
    pub async fn build(self) -> Result<Converter> {
        Converter::new(self.config).await
    }
}

impl Default for ConverterBuilder {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    // ========== ConverterBuilder tests ==========

    #[test]
    fn test_builder_default() {
        let builder = ConverterBuilder::new();
        
        // Check defaults match ConverterConfig::default()
        let default_config = ConverterConfig::default();
        assert_eq!(builder.config.pool.pool_size, default_config.pool.pool_size);
        assert_eq!(builder.config.render.dpi, default_config.render.dpi);
    }

    #[test]
    fn test_builder_pool_size() {
        let builder = ConverterBuilder::new().pool_size(8);
        assert_eq!(builder.config.pool.pool_size, 8);
    }

    #[test]
    fn test_builder_dpi() {
        let builder = ConverterBuilder::new().dpi(300);
        assert_eq!(builder.config.render.dpi, 300);
    }

    #[test]
    fn test_builder_render_threads() {
        let builder = ConverterBuilder::new().render_threads(4);
        assert_eq!(builder.config.render.render_threads, 4);
    }

    #[test]
    fn test_builder_conversion_timeout() {
        let builder = ConverterBuilder::new()
            .conversion_timeout(Duration::from_secs(120));
        assert_eq!(builder.config.pool.conversion_timeout, Duration::from_secs(120));
    }

    #[test]
    fn test_builder_soffice_path() {
        let path = PathBuf::from("/custom/path/to/soffice");
        let builder = ConverterBuilder::new().soffice_path(path.clone());
        assert_eq!(builder.config.pool.soffice_path, Some(path));
    }

    #[test]
    fn test_builder_temp_dir() {
        let path = PathBuf::from("/custom/temp");
        let builder = ConverterBuilder::new().temp_dir(path.clone());
        assert_eq!(builder.config.pool.temp_dir, Some(path));
    }

    #[test]
    fn test_builder_chaining() {
        let builder = ConverterBuilder::new()
            .pool_size(4)
            .dpi(150)
            .render_threads(2)
            .conversion_timeout(Duration::from_secs(90));

        assert_eq!(builder.config.pool.pool_size, 4);
        assert_eq!(builder.config.render.dpi, 150);
        assert_eq!(builder.config.render.render_threads, 2);
        assert_eq!(builder.config.pool.conversion_timeout, Duration::from_secs(90));
    }

    #[test]
    fn test_builder_default_trait() {
        let builder: ConverterBuilder = Default::default();
        assert_eq!(builder.config.pool.pool_size, ConverterConfig::default().pool.pool_size);
    }

    // ========== ConverterStats tests ==========

    #[test]
    fn test_converter_stats_struct() {
        let stats = ConverterStats {
            total_documents_processed: 42,
            pool_size: 4,
            dpi: 150,
        };

        assert_eq!(stats.total_documents_processed, 42);
        assert_eq!(stats.pool_size, 4);
        assert_eq!(stats.dpi, 150);
    }

    #[test]
    fn test_converter_stats_clone() {
        let stats = ConverterStats {
            total_documents_processed: 100,
            pool_size: 8,
            dpi: 300,
        };

        let cloned = stats.clone();
        assert_eq!(cloned.total_documents_processed, 100);
        assert_eq!(cloned.pool_size, 8);
        assert_eq!(cloned.dpi, 300);
    }

    #[test]
    fn test_converter_stats_debug() {
        let stats = ConverterStats {
            total_documents_processed: 50,
            pool_size: 2,
            dpi: 72,
        };

        let debug_str = format!("{:?}", stats);
        assert!(debug_str.contains("50"));
        assert!(debug_str.contains("2"));
        assert!(debug_str.contains("72"));
    }

    // ========== ConversionRequest tests ==========

    #[test]
    fn test_conversion_request_get_output_prefix() {
        let request = ConversionRequest {
            input_path: PathBuf::from("/path/to/document.docx"),
            output_dir: PathBuf::from("/output"),
            output_prefix: Some("custom_prefix".to_string()),
            dpi_override: None,
        };

        assert_eq!(request.get_output_prefix(), "custom_prefix");
    }

    #[test]
    fn test_conversion_request_get_output_prefix_from_filename() {
        let request = ConversionRequest {
            input_path: PathBuf::from("/path/to/my_document.docx"),
            output_dir: PathBuf::from("/output"),
            output_prefix: None,
            dpi_override: None,
        };

        assert_eq!(request.get_output_prefix(), "my_document");
    }

    #[test]
    fn test_conversion_request_dpi_override() {
        let request = ConversionRequest {
            input_path: PathBuf::from("/path/to/document.docx"),
            output_dir: PathBuf::from("/output"),
            output_prefix: None,
            dpi_override: Some(300),
        };

        assert_eq!(request.dpi_override, Some(300));
    }

    // ========== FileResult tests ==========

    #[test]
    fn test_file_result_struct() {
        let result = FileResult {
            input_path: PathBuf::from("/input/doc.docx"),
            output_paths: vec![
                PathBuf::from("/output/doc_page_0001.png"),
                PathBuf::from("/output/doc_page_0002.png"),
            ],
            page_count: 2,
            duration: Duration::from_millis(500),
        };

        assert_eq!(result.page_count, 2);
        assert_eq!(result.output_paths.len(), 2);
        assert_eq!(result.duration.as_millis(), 500);
    }

    // ========== BatchResult tests ==========

    #[test]
    fn test_batch_result_empty() {
        let result = BatchResult {
            successful: vec![],
            failed: vec![],
            total_duration: Duration::from_secs(0),
            total_pages: 0,
        };

        assert!(result.successful.is_empty());
        assert!(result.failed.is_empty());
        assert_eq!(result.total_pages, 0);
    }

    #[test]
    fn test_batch_result_with_data() {
        let result = BatchResult {
            successful: vec![
                FileResult {
                    input_path: PathBuf::from("/input/doc1.docx"),
                    output_paths: vec![PathBuf::from("/output/doc1_page_0001.png")],
                    page_count: 1,
                    duration: Duration::from_millis(100),
                },
            ],
            failed: vec![
                FailedFile {
                    input_path: PathBuf::from("/input/bad.docx"),
                    error: "Conversion failed".to_string(),
                },
            ],
            total_duration: Duration::from_secs(1),
            total_pages: 1,
        };

        assert_eq!(result.successful.len(), 1);
        assert_eq!(result.failed.len(), 1);
        assert_eq!(result.total_pages, 1);
    }

    // ========== ConversionProgress tests ==========

    #[test]
    fn test_conversion_progress_struct() {
        let progress = ConversionProgress {
            file_index: 2,
            total_files: 10,
            current_file: "document.docx".to_string(),
            pages_completed: 5,
            total_pages: Some(10),
            stage: ConversionStage::RenderingPages,
        };

        assert_eq!(progress.file_index, 2);
        assert_eq!(progress.total_files, 10);
        assert_eq!(progress.pages_completed, 5);
        assert_eq!(progress.total_pages, Some(10));
    }

    #[test]
    fn test_conversion_stage_variants() {
        // Test all stage variants exist
        let stages = [
            ConversionStage::ConvertingToPdf,
            ConversionStage::RenderingPages,
            ConversionStage::Completed,
            ConversionStage::Failed,
        ];

        for stage in &stages {
            let _ = format!("{:?}", stage); // Should not panic
        }
    }

    // ========== Converter integration tests (require LibreOffice + pdfium) ==========

    #[tokio::test]
    async fn test_converter_creation_fails_without_libreoffice() {
        // With a non-existent soffice path
        let config = ConverterConfig {
            pool: crate::config::PoolConfig {
                soffice_path: Some(PathBuf::from("/nonexistent/soffice")),
                ..Default::default()
            },
            render: RenderConfig::default(),
        };

        let result = Converter::new(config).await;
        assert!(result.is_err());
    }

    #[tokio::test]
    async fn test_converter_builder_build_fails_without_libreoffice() {
        let result = ConverterBuilder::new()
            .soffice_path(PathBuf::from("/nonexistent/soffice"))
            .build()
            .await;

        assert!(result.is_err());
    }
}
