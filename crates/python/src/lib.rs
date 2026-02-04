//! Python bindings for office-to-png using PyO3.
//!
//! This module provides a Python interface to the high-performance
//! Office-to-PNG conversion library.
//!
//! # Example
//!
//! ```python
//! import asyncio
//! from office_to_png import OfficeConverter, ConversionRequest
//!
//! async def main():
//!     # Create converter with 4 workers and 300 DPI
//!     converter = OfficeConverter(pool_size=4, dpi=300)
//!     
//!     # Convert a single file
//!     result = await converter.convert("document.docx", "./output")
//!     print(f"Rendered {result.page_count} pages")
//!     
//!     # Batch conversion with progress
//!     def on_progress(p):
//!         print(f"File {p.file_index + 1}/{p.total_files}: {p.pages_completed} pages")
//!     
//!     results = await converter.convert_batch(
//!         ["doc1.docx", "doc2.xlsx"],
//!         "./output",
//!         progress_callback=on_progress
//!     )
//!
//! asyncio.run(main())
//! ```

use office_to_png_core::{
    BatchResult, ConversionProgress, ConversionRequest, ConversionStage, Converter,
    ConverterBuilder, ConverterConfig, FileResult, PngPage,
};
use pyo3::exceptions::{PyIOError, PyRuntimeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};
use std::path::PathBuf;
use std::sync::Arc;
use tokio::runtime::Runtime;
use tokio::sync::Mutex;

/// Python wrapper for ConversionProgress.
#[pyclass(name = "ConversionProgress")]
#[derive(Clone)]
pub struct PyConversionProgress {
    #[pyo3(get)]
    pub file_index: usize,
    #[pyo3(get)]
    pub total_files: usize,
    #[pyo3(get)]
    pub current_file: String,
    #[pyo3(get)]
    pub pages_completed: usize,
    #[pyo3(get)]
    pub total_pages: Option<usize>,
    #[pyo3(get)]
    pub stage: String,
}

impl From<ConversionProgress> for PyConversionProgress {
    fn from(p: ConversionProgress) -> Self {
        let stage = match p.stage {
            ConversionStage::Queued => "queued",
            ConversionStage::ConvertingToPdf => "converting_to_pdf",
            ConversionStage::RenderingPages => "rendering_pages",
            ConversionStage::EncodingPng => "encoding_png",
            ConversionStage::Completed => "completed",
            ConversionStage::Failed => "failed",
        };

        Self {
            file_index: p.file_index,
            total_files: p.total_files,
            current_file: p.current_file,
            pages_completed: p.pages_completed,
            total_pages: p.total_pages,
            stage: stage.to_string(),
        }
    }
}

#[pymethods]
impl PyConversionProgress {
    fn __repr__(&self) -> String {
        format!(
            "ConversionProgress(file={}/{}, current='{}', pages={}/{}, stage='{}')",
            self.file_index + 1,
            self.total_files,
            self.current_file,
            self.pages_completed,
            self.total_pages
                .map(|n| n.to_string())
                .unwrap_or_else(|| "?".to_string()),
            self.stage
        )
    }
}

/// Python wrapper for FileResult.
#[pyclass(name = "FileResult")]
#[derive(Clone)]
pub struct PyFileResult {
    #[pyo3(get)]
    pub input_path: String,
    #[pyo3(get)]
    pub output_paths: Vec<String>,
    #[pyo3(get)]
    pub page_count: usize,
    #[pyo3(get)]
    pub duration_secs: f64,
}

impl From<FileResult> for PyFileResult {
    fn from(r: FileResult) -> Self {
        Self {
            input_path: r.input_path.to_string_lossy().to_string(),
            output_paths: r
                .output_paths
                .iter()
                .map(|p| p.to_string_lossy().to_string())
                .collect(),
            page_count: r.page_count,
            duration_secs: r.duration.as_secs_f64(),
        }
    }
}

#[pymethods]
impl PyFileResult {
    fn __repr__(&self) -> String {
        format!(
            "FileResult(input='{}', pages={}, duration={:.2}s)",
            self.input_path, self.page_count, self.duration_secs
        )
    }
}

/// Python wrapper for BatchResult.
#[pyclass(name = "BatchResult")]
#[derive(Clone)]
pub struct PyBatchResult {
    #[pyo3(get)]
    pub successful: Vec<PyFileResult>,
    #[pyo3(get)]
    pub failed: Vec<(String, String)>, // (path, error)
    #[pyo3(get)]
    pub total_duration_secs: f64,
    #[pyo3(get)]
    pub total_pages: usize,
}

impl From<BatchResult> for PyBatchResult {
    fn from(r: BatchResult) -> Self {
        Self {
            successful: r.successful.into_iter().map(PyFileResult::from).collect(),
            failed: r
                .failed
                .into_iter()
                .map(|f| (f.input_path.to_string_lossy().to_string(), f.error))
                .collect(),
            total_duration_secs: r.total_duration.as_secs_f64(),
            total_pages: r.total_pages,
        }
    }
}

#[pymethods]
impl PyBatchResult {
    fn __repr__(&self) -> String {
        format!(
            "BatchResult(successful={}, failed={}, pages={}, duration={:.2}s)",
            self.successful.len(),
            self.failed.len(),
            self.total_pages,
            self.total_duration_secs
        )
    }

    /// Get the number of successful conversions.
    #[getter]
    fn success_count(&self) -> usize {
        self.successful.len()
    }

    /// Get the number of failed conversions.
    #[getter]
    fn failure_count(&self) -> usize {
        self.failed.len()
    }

    /// Check if all conversions succeeded.
    #[getter]
    fn all_succeeded(&self) -> bool {
        self.failed.is_empty()
    }
}

/// Python wrapper for PngPage (for streaming).
#[pyclass(name = "PngPage")]
#[derive(Clone)]
pub struct PyPngPage {
    #[pyo3(get)]
    pub page_number: usize,
    #[pyo3(get)]
    pub width: u32,
    #[pyo3(get)]
    pub height: u32,
    #[pyo3(get)]
    pub output_path: Option<String>,
    data: Vec<u8>,
}

impl From<PngPage> for PyPngPage {
    fn from(p: PngPage) -> Self {
        Self {
            page_number: p.page_number,
            width: p.width,
            height: p.height,
            output_path: p.output_path.map(|p| p.to_string_lossy().to_string()),
            data: p.data,
        }
    }
}

#[pymethods]
impl PyPngPage {
    fn __repr__(&self) -> String {
        format!(
            "PngPage(number={}, size={}x{}, path={})",
            self.page_number,
            self.width,
            self.height,
            self.output_path
                .as_ref()
                .map(|s| format!("'{}'", s))
                .unwrap_or_else(|| "None".to_string())
        )
    }

    /// Get the raw PNG data as bytes.
    fn data(&self, py: Python<'_>) -> PyResult<Py<pyo3::types::PyBytes>> {
        Ok(pyo3::types::PyBytes::new(py, &self.data).into())
    }

    /// Get the size of the PNG data in bytes.
    #[getter]
    fn data_size(&self) -> usize {
        self.data.len()
    }
}

/// High-performance Office document to PNG converter.
///
/// This converter uses LibreOffice for document rendering and pdfium for
/// PDF-to-PNG conversion, with parallel processing for maximum throughput.
///
/// Args:
///     pool_size: Number of LibreOffice instances (default: CPU count)
///     dpi: Output DPI (default: 300)
///     conversion_timeout: Timeout per document in seconds (default: 120)
///     render_threads: Number of PNG rendering threads (default: CPU count)
///
/// Example:
///     >>> converter = OfficeConverter(pool_size=4, dpi=300)
///     >>> result = await converter.convert("doc.docx", "./output")
#[pyclass(name = "OfficeConverter")]
pub struct PyOfficeConverter {
    converter: Arc<Mutex<Option<Converter>>>,
    runtime: Arc<Runtime>,
    config: ConverterConfig,
}

#[pymethods]
impl PyOfficeConverter {
    #[new]
    #[pyo3(signature = (pool_size=None, dpi=None, conversion_timeout=None, render_threads=None))]
    fn new(
        pool_size: Option<usize>,
        dpi: Option<u32>,
        conversion_timeout: Option<u64>,
        render_threads: Option<usize>,
    ) -> PyResult<Self> {
        // Build config
        let mut config = ConverterConfig::default();
        if let Some(size) = pool_size {
            config.pool.pool_size = size;
        }
        if let Some(d) = dpi {
            config.render.dpi = d;
        }
        if let Some(timeout) = conversion_timeout {
            config.pool.conversion_timeout = std::time::Duration::from_secs(timeout);
        }
        if let Some(threads) = render_threads {
            config.render.render_threads = threads;
        }

        // Validate config
        config
            .validate()
            .map_err(|e| PyValueError::new_err(e.to_string()))?;

        // Create runtime
        let runtime = Runtime::new().map_err(|e| PyRuntimeError::new_err(e.to_string()))?;

        Ok(Self {
            converter: Arc::new(Mutex::new(None)),
            runtime: Arc::new(runtime),
            config,
        })
    }

    /// Initialize the converter (called automatically on first use).
    fn initialize<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyAny>> {
        let converter_arc = Arc::clone(&self.converter);
        let config = self.config.clone();
        let runtime = Arc::clone(&self.runtime);

        pyo3_async_runtimes::tokio::future_into_py(py, async move {
            let mut guard = converter_arc.lock().await;
            if guard.is_none() {
                let converter = Converter::new(config)
                    .await
                    .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
                *guard = Some(converter);
            }
            Ok(())
        })
    }

    /// Convert a single document to PNG images.
    ///
    /// Args:
    ///     input_path: Path to the input Office document
    ///     output_dir: Directory to write output PNGs
    ///     dpi: Optional DPI override for this conversion
    ///     output_prefix: Optional prefix for output filenames
    ///
    /// Returns:
    ///     FileResult with information about the conversion
    #[pyo3(signature = (input_path, output_dir, dpi=None, output_prefix=None))]
    fn convert<'py>(
        &self,
        py: Python<'py>,
        input_path: String,
        output_dir: String,
        dpi: Option<u32>,
        output_prefix: Option<String>,
    ) -> PyResult<Bound<'py, PyAny>> {
        let converter_arc = Arc::clone(&self.converter);
        let config = self.config.clone();
        let runtime = Arc::clone(&self.runtime);

        pyo3_async_runtimes::tokio::future_into_py(py, async move {
            // Ensure initialized
            let mut guard = converter_arc.lock().await;
            if guard.is_none() {
                let converter = Converter::new(config)
                    .await
                    .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
                *guard = Some(converter);
            }

            let converter = guard.as_ref().unwrap();

            // Build request
            let mut request = ConversionRequest::new(input_path, output_dir);
            if let Some(d) = dpi {
                request = request.with_dpi(d);
            }
            if let Some(prefix) = output_prefix {
                request = request.with_prefix(prefix);
            }

            // Convert
            let result = converter
                .convert(request)
                .await
                .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;

            Ok(PyFileResult::from(result))
        })
    }

    /// Convert multiple documents in batch.
    ///
    /// Args:
    ///     input_paths: List of paths to input documents
    ///     output_dir: Directory to write output PNGs
    ///     dpi: Optional DPI override for all conversions
    ///     progress_callback: Optional callback for progress updates
    ///     concurrency: Number of documents to process in parallel (default: pool_size)
    ///
    /// Returns:
    ///     BatchResult with information about all conversions
    #[pyo3(signature = (input_paths, output_dir, dpi=None, progress_callback=None, concurrency=None))]
    fn convert_batch<'py>(
        &self,
        py: Python<'py>,
        input_paths: Vec<String>,
        output_dir: String,
        dpi: Option<u32>,
        progress_callback: Option<PyObject>,
        concurrency: Option<usize>,
    ) -> PyResult<Bound<'py, PyAny>> {
        let converter_arc = Arc::clone(&self.converter);
        let config = self.config.clone();
        let pool_size = self.config.pool.pool_size;

        pyo3_async_runtimes::tokio::future_into_py(py, async move {
            // Ensure initialized
            let mut guard = converter_arc.lock().await;
            if guard.is_none() {
                let converter = Converter::new(config)
                    .await
                    .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
                *guard = Some(converter);
            }

            let converter = guard.as_ref().unwrap();

            // Build requests
            let requests: Vec<ConversionRequest> = input_paths
                .into_iter()
                .map(|path| {
                    let mut req = ConversionRequest::new(path, output_dir.clone());
                    if let Some(d) = dpi {
                        req = req.with_dpi(d);
                    }
                    req
                })
                .collect();

            // Convert with or without progress
            let result = if let Some(callback) = progress_callback {
                converter
                    .convert_batch_with_progress(requests, move |progress| {
                        Python::with_gil(|py| {
                            let py_progress = PyConversionProgress::from(progress);
                            if let Err(e) = callback.call1(py, (py_progress,)) {
                                eprintln!("Progress callback error: {}", e);
                            }
                        });
                    })
                    .await
            } else {
                let concurrent = concurrency.unwrap_or(pool_size);
                converter.convert_parallel(requests, concurrent).await
            };

            Ok(PyBatchResult::from(result))
        })
    }

    /// Get health information about the converter.
    fn health<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyAny>> {
        let converter_arc = Arc::clone(&self.converter);

        pyo3_async_runtimes::tokio::future_into_py(py, async move {
            let guard = converter_arc.lock().await;
            if let Some(converter) = guard.as_ref() {
                let health = converter.health().await;
                Ok(format!(
                    "Pool: {}/{} instances, {} total processed, shutdown={}",
                    health
                        .instances
                        .iter()
                        .filter(|i| !i.is_busy)
                        .count(),
                    health.pool_size,
                    health.total_processed,
                    health.is_shutdown
                ))
            } else {
                Ok("Not initialized".to_string())
            }
        })
    }

    /// Shutdown the converter and release resources.
    fn shutdown<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyAny>> {
        let converter_arc = Arc::clone(&self.converter);

        pyo3_async_runtimes::tokio::future_into_py(py, async move {
            let mut guard = converter_arc.lock().await;
            if let Some(converter) = guard.take() {
                converter.shutdown().await;
            }
            Ok(())
        })
    }

    /// Get the configured pool size.
    #[getter]
    fn pool_size(&self) -> usize {
        self.config.pool.pool_size
    }

    /// Get the configured DPI.
    #[getter]
    fn dpi(&self) -> u32 {
        self.config.render.dpi
    }

    fn __repr__(&self) -> String {
        format!(
            "OfficeConverter(pool_size={}, dpi={})",
            self.config.pool.pool_size, self.config.render.dpi
        )
    }
}

/// Iterator for streaming page-by-page conversion.
#[pyclass(name = "PageIterator")]
pub struct PyPageIterator {
    pages: Vec<PyPngPage>,
    current: usize,
}

#[pymethods]
impl PyPageIterator {
    fn __iter__(slf: PyRef<Self>) -> PyRef<Self> {
        slf
    }

    fn __next__(&mut self) -> Option<PyPngPage> {
        if self.current < self.pages.len() {
            let page = self.pages[self.current].clone();
            self.current += 1;
            Some(page)
        } else {
            None
        }
    }

    fn __len__(&self) -> usize {
        self.pages.len()
    }
}

/// Check if LibreOffice is installed and available.
#[pyfunction]
fn is_libreoffice_available() -> bool {
    which::which("soffice")
        .or_else(|_| which::which("libreoffice"))
        .is_ok()
        || std::path::Path::new("/Applications/LibreOffice.app/Contents/MacOS/soffice").exists()
        || std::path::Path::new("/usr/bin/soffice").exists()
}

/// Get the path to the LibreOffice binary if available.
#[pyfunction]
fn get_libreoffice_path() -> Option<String> {
    // Check common locations
    let candidates = [
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/usr/bin/soffice",
        "/usr/lib/libreoffice/program/soffice",
        "/opt/libreoffice/program/soffice",
        "/snap/bin/libreoffice.soffice",
    ];

    for candidate in candidates {
        if std::path::Path::new(candidate).exists() {
            return Some(candidate.to_string());
        }
    }

    // Try PATH
    which::which("soffice")
        .or_else(|_| which::which("libreoffice"))
        .ok()
        .map(|p| p.to_string_lossy().to_string())
}

/// Get the list of supported file extensions.
#[pyfunction]
fn supported_extensions() -> Vec<&'static str> {
    office_to_png_core::SUPPORTED_EXTENSIONS.to_vec()
}

/// Check if a file extension is supported.
#[pyfunction]
fn is_supported_extension(ext: &str) -> bool {
    office_to_png_core::is_supported_extension(ext)
}

/// Initialize logging for the library.
#[pyfunction]
fn init_logging() {
    office_to_png_core::init_logging();
}

/// Python module definition.
#[pymodule]
fn office_to_png(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<PyOfficeConverter>()?;
    m.add_class::<PyConversionProgress>()?;
    m.add_class::<PyFileResult>()?;
    m.add_class::<PyBatchResult>()?;
    m.add_class::<PyPngPage>()?;
    m.add_class::<PyPageIterator>()?;

    m.add_function(wrap_pyfunction!(is_libreoffice_available, m)?)?;
    m.add_function(wrap_pyfunction!(get_libreoffice_path, m)?)?;
    m.add_function(wrap_pyfunction!(supported_extensions, m)?)?;
    m.add_function(wrap_pyfunction!(is_supported_extension, m)?)?;
    m.add_function(wrap_pyfunction!(init_logging, m)?)?;

    Ok(())
}
