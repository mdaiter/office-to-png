//! Configuration types for office-to-png conversion.

use serde::{Deserialize, Serialize};
use std::path::PathBuf;
use std::time::Duration;

/// Configuration for the LibreOffice process pool.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PoolConfig {
    /// Number of LibreOffice instances in the pool.
    /// Default: number of CPU cores.
    pub pool_size: usize,

    /// Timeout for individual document conversions.
    /// Default: 120 seconds.
    pub conversion_timeout: Duration,

    /// Maximum number of documents to process per instance before recycling.
    /// This helps prevent memory leaks in LibreOffice.
    /// Default: 100 documents.
    pub max_docs_per_instance: u32,

    /// Base port for UNO socket connections (if using persistent mode).
    /// Each instance uses base_port + instance_index.
    /// Default: 2002.
    pub base_port: u16,

    /// Directory for temporary files.
    /// Default: system temp directory.
    pub temp_dir: Option<PathBuf>,

    /// Path to soffice binary. If None, searches PATH.
    pub soffice_path: Option<PathBuf>,

    /// Whether to use persistent LibreOffice instances (faster but more complex).
    /// Default: false (use one-shot conversion mode).
    pub use_persistent_instances: bool,

    /// Time to wait for instance startup.
    /// Default: 30 seconds.
    pub instance_startup_timeout: Duration,
}

impl Default for PoolConfig {
    fn default() -> Self {
        Self {
            pool_size: num_cpus::get(),
            conversion_timeout: Duration::from_secs(120),
            max_docs_per_instance: 100,
            base_port: 2002,
            temp_dir: None,
            soffice_path: None,
            use_persistent_instances: false,
            instance_startup_timeout: Duration::from_secs(30),
        }
    }
}

impl PoolConfig {
    /// Create a new pool config with specified pool size.
    pub fn with_pool_size(pool_size: usize) -> Self {
        Self {
            pool_size,
            ..Default::default()
        }
    }

    /// Set the conversion timeout.
    pub fn conversion_timeout(mut self, timeout: Duration) -> Self {
        self.conversion_timeout = timeout;
        self
    }

    /// Set the maximum documents per instance before recycling.
    pub fn max_docs_per_instance(mut self, max: u32) -> Self {
        self.max_docs_per_instance = max;
        self
    }

    /// Set the temporary directory.
    pub fn temp_dir(mut self, dir: PathBuf) -> Self {
        self.temp_dir = Some(dir);
        self
    }

    /// Set the soffice binary path.
    pub fn soffice_path(mut self, path: PathBuf) -> Self {
        self.soffice_path = Some(path);
        self
    }

    /// Enable persistent instances for better performance.
    pub fn use_persistent_instances(mut self, enabled: bool) -> Self {
        self.use_persistent_instances = enabled;
        self
    }

    /// Validate the configuration.
    pub fn validate(&self) -> crate::error::Result<()> {
        if self.pool_size == 0 {
            return Err(crate::error::ConversionError::InvalidConfig(
                "pool_size must be at least 1".to_string(),
            ));
        }
        if self.conversion_timeout.as_secs() == 0 {
            return Err(crate::error::ConversionError::InvalidConfig(
                "conversion_timeout must be greater than 0".to_string(),
            ));
        }
        Ok(())
    }
}

/// Configuration for PDF to PNG rendering.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RenderConfig {
    /// Output DPI (dots per inch).
    /// Default: 300.
    pub dpi: u32,

    /// Number of threads for parallel page rendering.
    /// Default: number of CPU cores.
    pub render_threads: usize,

    /// PNG compression level (0-9, higher = smaller file, slower).
    /// Default: 6.
    pub png_compression: u8,

    /// Whether to use alpha channel (transparency).
    /// Default: false.
    pub use_alpha: bool,

    /// Background color for pages (if not using alpha).
    /// Default: white (255, 255, 255).
    pub background_color: (u8, u8, u8),
}

impl Default for RenderConfig {
    fn default() -> Self {
        Self {
            dpi: 300,
            render_threads: num_cpus::get(),
            png_compression: 6,
            use_alpha: false,
            background_color: (255, 255, 255),
        }
    }
}

impl RenderConfig {
    /// Create a render config with specified DPI.
    pub fn with_dpi(dpi: u32) -> Self {
        Self {
            dpi,
            ..Default::default()
        }
    }

    /// Set the number of render threads.
    pub fn render_threads(mut self, threads: usize) -> Self {
        self.render_threads = threads;
        self
    }

    /// Set PNG compression level.
    pub fn png_compression(mut self, level: u8) -> Self {
        self.png_compression = level.min(9);
        self
    }

    /// Enable alpha channel.
    pub fn use_alpha(mut self, enabled: bool) -> Self {
        self.use_alpha = enabled;
        self
    }

    /// Validate the configuration.
    pub fn validate(&self) -> crate::error::Result<()> {
        if self.dpi == 0 || self.dpi > 1200 {
            return Err(crate::error::ConversionError::InvalidConfig(
                "dpi must be between 1 and 1200".to_string(),
            ));
        }
        if self.render_threads == 0 {
            return Err(crate::error::ConversionError::InvalidConfig(
                "render_threads must be at least 1".to_string(),
            ));
        }
        Ok(())
    }
}

/// Combined configuration for the converter.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ConverterConfig {
    /// Pool configuration.
    pub pool: PoolConfig,

    /// Render configuration.
    pub render: RenderConfig,
}

impl Default for ConverterConfig {
    fn default() -> Self {
        Self {
            pool: PoolConfig::default(),
            render: RenderConfig::default(),
        }
    }
}

impl ConverterConfig {
    /// Create a new converter config with specified pool size and DPI.
    pub fn new(pool_size: usize, dpi: u32) -> Self {
        Self {
            pool: PoolConfig::with_pool_size(pool_size),
            render: RenderConfig::with_dpi(dpi),
        }
    }

    /// Validate the entire configuration.
    pub fn validate(&self) -> crate::error::Result<()> {
        self.pool.validate()?;
        self.render.validate()?;
        Ok(())
    }
}

/// A single conversion request.
#[derive(Debug, Clone)]
pub struct ConversionRequest {
    /// Path to the input Office document.
    pub input_path: PathBuf,

    /// Directory to write output PNGs.
    pub output_dir: PathBuf,

    /// Optional prefix for output filenames.
    /// Default: input filename without extension.
    pub output_prefix: Option<String>,

    /// Override DPI for this specific conversion.
    pub dpi_override: Option<u32>,
}

impl ConversionRequest {
    /// Create a new conversion request.
    pub fn new(input_path: impl Into<PathBuf>, output_dir: impl Into<PathBuf>) -> Self {
        Self {
            input_path: input_path.into(),
            output_dir: output_dir.into(),
            output_prefix: None,
            dpi_override: None,
        }
    }

    /// Set a custom output prefix.
    pub fn with_prefix(mut self, prefix: impl Into<String>) -> Self {
        self.output_prefix = Some(prefix.into());
        self
    }

    /// Override DPI for this conversion.
    pub fn with_dpi(mut self, dpi: u32) -> Self {
        self.dpi_override = Some(dpi);
        self
    }

    /// Get the output prefix, defaulting to the input filename.
    pub fn get_output_prefix(&self) -> String {
        self.output_prefix.clone().unwrap_or_else(|| {
            self.input_path
                .file_stem()
                .and_then(|s| s.to_str())
                .unwrap_or("output")
                .to_string()
        })
    }
}

/// Progress information for a conversion operation.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ConversionProgress {
    /// Index of the current file being processed.
    pub file_index: usize,

    /// Total number of files to process.
    pub total_files: usize,

    /// Name of the current file.
    pub current_file: String,

    /// Number of pages completed for the current file.
    pub pages_completed: usize,

    /// Total pages in the current file (if known).
    pub total_pages: Option<usize>,

    /// Current stage of processing.
    pub stage: ConversionStage,
}

/// Stage of the conversion process.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum ConversionStage {
    /// Queued, waiting to start.
    Queued,
    /// Converting to PDF via LibreOffice.
    ConvertingToPdf,
    /// Rendering PDF pages to PNG.
    RenderingPages,
    /// Encoding PNG files.
    EncodingPng,
    /// Completed successfully.
    Completed,
    /// Failed with error.
    Failed,
}

/// Result of a batch conversion operation.
#[derive(Debug, Clone)]
pub struct BatchResult {
    /// Successfully converted files.
    pub successful: Vec<FileResult>,

    /// Failed conversions.
    pub failed: Vec<FailedFile>,

    /// Total processing time.
    pub total_duration: Duration,

    /// Total pages rendered.
    pub total_pages: usize,
}

/// Result for a single successfully converted file.
#[derive(Debug, Clone)]
pub struct FileResult {
    /// Original input path.
    pub input_path: PathBuf,

    /// Output PNG paths.
    pub output_paths: Vec<PathBuf>,

    /// Number of pages.
    pub page_count: usize,

    /// Processing time for this file.
    pub duration: Duration,
}

/// Information about a failed conversion.
#[derive(Debug, Clone)]
pub struct FailedFile {
    /// Original input path.
    pub input_path: PathBuf,

    /// Error message.
    pub error: String,
}

/// A single rendered page.
#[derive(Debug, Clone)]
pub struct PngPage {
    /// Page number (1-indexed).
    pub page_number: usize,

    /// PNG image data.
    pub data: Vec<u8>,

    /// Image width in pixels.
    pub width: u32,

    /// Image height in pixels.
    pub height: u32,

    /// Path where the PNG was written (if saved to disk).
    pub output_path: Option<PathBuf>,
}

#[cfg(test)]
mod tests {
    use super::*;

    // PoolConfig tests
    #[test]
    fn test_pool_config_defaults() {
        let config = PoolConfig::default();
        assert!(config.pool_size > 0);
        assert_eq!(config.conversion_timeout.as_secs(), 120);
        assert_eq!(config.max_docs_per_instance, 100);
        assert_eq!(config.base_port, 2002);
        assert!(config.temp_dir.is_none());
        assert!(config.soffice_path.is_none());
        assert!(!config.use_persistent_instances);
    }

    #[test]
    fn test_pool_config_with_pool_size() {
        let config = PoolConfig::with_pool_size(8);
        assert_eq!(config.pool_size, 8);
    }

    #[test]
    fn test_pool_config_builder_pattern() {
        let config = PoolConfig::with_pool_size(4)
            .conversion_timeout(Duration::from_secs(60))
            .max_docs_per_instance(50)
            .use_persistent_instances(true);

        assert_eq!(config.pool_size, 4);
        assert_eq!(config.conversion_timeout.as_secs(), 60);
        assert_eq!(config.max_docs_per_instance, 50);
        assert!(config.use_persistent_instances);
    }

    #[test]
    fn test_pool_config_validation_valid() {
        let config = PoolConfig::with_pool_size(4);
        assert!(config.validate().is_ok());
    }

    #[test]
    fn test_pool_config_validation_zero_pool_size() {
        let mut config = PoolConfig::default();
        config.pool_size = 0;
        assert!(config.validate().is_err());
    }

    #[test]
    fn test_pool_config_validation_zero_timeout() {
        let mut config = PoolConfig::default();
        config.conversion_timeout = Duration::from_secs(0);
        assert!(config.validate().is_err());
    }

    // RenderConfig tests
    #[test]
    fn test_render_config_defaults() {
        let config = RenderConfig::default();
        assert_eq!(config.dpi, 300);
        assert!(config.render_threads > 0);
        assert_eq!(config.png_compression, 6);
        assert!(!config.use_alpha);
        assert_eq!(config.background_color, (255, 255, 255));
    }

    #[test]
    fn test_render_config_with_dpi() {
        let config = RenderConfig::with_dpi(150);
        assert_eq!(config.dpi, 150);
    }

    #[test]
    fn test_render_config_builder_pattern() {
        let config = RenderConfig::with_dpi(72)
            .render_threads(2)
            .png_compression(9)
            .use_alpha(true);

        assert_eq!(config.dpi, 72);
        assert_eq!(config.render_threads, 2);
        assert_eq!(config.png_compression, 9);
        assert!(config.use_alpha);
    }

    #[test]
    fn test_render_config_png_compression_clamped() {
        let config = RenderConfig::default().png_compression(15);
        assert_eq!(config.png_compression, 9); // Clamped to max 9
    }

    #[test]
    fn test_render_config_validation_valid() {
        let config = RenderConfig::with_dpi(300);
        assert!(config.validate().is_ok());
    }

    #[test]
    fn test_render_config_validation_zero_dpi() {
        let mut config = RenderConfig::default();
        config.dpi = 0;
        assert!(config.validate().is_err());
    }

    #[test]
    fn test_render_config_validation_excessive_dpi() {
        let mut config = RenderConfig::default();
        config.dpi = 1201;
        assert!(config.validate().is_err());
    }

    #[test]
    fn test_render_config_validation_zero_threads() {
        let mut config = RenderConfig::default();
        config.render_threads = 0;
        assert!(config.validate().is_err());
    }

    // ConverterConfig tests
    #[test]
    fn test_converter_config_new() {
        let config = ConverterConfig::new(4, 150);
        assert_eq!(config.pool.pool_size, 4);
        assert_eq!(config.render.dpi, 150);
    }

    #[test]
    fn test_converter_config_validate_propagates() {
        let mut config = ConverterConfig::default();
        config.pool.pool_size = 0;
        assert!(config.validate().is_err());

        let mut config2 = ConverterConfig::default();
        config2.render.dpi = 0;
        assert!(config2.validate().is_err());
    }

    // ConversionRequest tests
    #[test]
    fn test_conversion_request_new() {
        let request = ConversionRequest::new("input.docx", "/output");
        assert_eq!(request.input_path, PathBuf::from("input.docx"));
        assert_eq!(request.output_dir, PathBuf::from("/output"));
        assert!(request.output_prefix.is_none());
        assert!(request.dpi_override.is_none());
    }

    #[test]
    fn test_conversion_request_with_prefix() {
        let request = ConversionRequest::new("input.docx", "/output").with_prefix("custom_prefix");
        assert_eq!(request.output_prefix, Some("custom_prefix".to_string()));
    }

    #[test]
    fn test_conversion_request_with_dpi() {
        let request = ConversionRequest::new("input.docx", "/output").with_dpi(150);
        assert_eq!(request.dpi_override, Some(150));
    }

    #[test]
    fn test_conversion_request_get_output_prefix_custom() {
        let request = ConversionRequest::new("input.docx", "/output").with_prefix("my_doc");
        assert_eq!(request.get_output_prefix(), "my_doc");
    }

    #[test]
    fn test_conversion_request_get_output_prefix_from_filename() {
        let request = ConversionRequest::new("path/to/document.docx", "/output");
        assert_eq!(request.get_output_prefix(), "document");
    }

    #[test]
    fn test_conversion_request_get_output_prefix_no_extension() {
        let request = ConversionRequest::new("path/to/document", "/output");
        assert_eq!(request.get_output_prefix(), "document");
    }

    // ConversionStage tests
    #[test]
    fn test_conversion_stage_variants() {
        assert_ne!(ConversionStage::Queued, ConversionStage::Completed);
        assert_eq!(ConversionStage::Failed, ConversionStage::Failed);
    }

    // PngPage tests
    #[test]
    fn test_png_page_creation() {
        let page = PngPage {
            page_number: 1,
            data: vec![0x89, 0x50, 0x4E, 0x47], // PNG magic bytes
            width: 100,
            height: 200,
            output_path: None,
        };
        assert_eq!(page.page_number, 1);
        assert_eq!(page.width, 100);
        assert_eq!(page.height, 200);
    }
}
