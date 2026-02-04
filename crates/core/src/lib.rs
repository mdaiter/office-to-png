//! # office-to-png-core
//!
//! High-performance Office document to PNG conversion library.
//!
//! This library provides a fast, parallelized pipeline for converting
//! Microsoft Office documents (.docx, .xlsx) to PNG images using:
//!
//! - **LibreOffice** for document-to-PDF conversion
//! - **pdfium** (Google's PDF engine) for PDF rendering
//! - **SIMD-accelerated** PNG encoding
//! - **Parallel processing** via tokio and rayon
//!
//! ## Quick Start
//!
//! ```rust,no_run
//! use office_to_png_core::{Converter, ConverterConfig, ConversionRequest};
//!
//! #[tokio::main]
//! async fn main() -> anyhow::Result<()> {
//!     // Create converter with 4 LibreOffice instances and 300 DPI
//!     let config = ConverterConfig::new(4, 300);
//!     let converter = Converter::new(config).await?;
//!
//!     // Convert a document
//!     let request = ConversionRequest::new("document.docx", "./output");
//!     let result = converter.convert(request).await?;
//!
//!     println!("Rendered {} pages", result.page_count);
//!     Ok(())
//! }
//! ```
//!
//! ## Batch Processing with Progress
//!
//! ```rust,no_run
//! use office_to_png_core::{Converter, ConverterConfig, ConversionRequest, ConversionProgress};
//!
//! #[tokio::main]
//! async fn main() -> anyhow::Result<()> {
//!     let converter = Converter::new(ConverterConfig::default()).await?;
//!
//!     let requests = vec![
//!         ConversionRequest::new("doc1.docx", "./output"),
//!         ConversionRequest::new("doc2.xlsx", "./output"),
//!     ];
//!
//!     let result = converter.convert_batch_with_progress(
//!         requests,
//!         |progress: ConversionProgress| {
//!             println!("File {}/{}: {} pages done",
//!                 progress.file_index + 1,
//!                 progress.total_files,
//!                 progress.pages_completed
//!             );
//!         }
//!     ).await;
//!
//!     println!("Total: {} pages in {:?}", result.total_pages, result.total_duration);
//!     Ok(())
//! }
//! ```

pub mod config;
pub mod converter;
pub mod error;
pub mod pdf_renderer;
pub mod pool;

// Re-export main types for convenience
pub use config::{
    BatchResult, ConversionProgress, ConversionRequest, ConversionStage, ConverterConfig,
    FailedFile, FileResult, PngPage, PoolConfig, RenderConfig,
};
pub use converter::Converter;
pub use error::{ConversionError, Result};
pub use pdf_renderer::PdfRenderer;
pub use pool::LibreOfficePool;

/// Supported Office file extensions.
pub const SUPPORTED_EXTENSIONS: &[&str] = &["docx", "doc", "xlsx", "xls"];

/// Check if a file extension is supported.
pub fn is_supported_extension(ext: &str) -> bool {
    SUPPORTED_EXTENSIONS
        .iter()
        .any(|&e| e.eq_ignore_ascii_case(ext))
}

/// Initialize the library's logging.
/// Call this once at application startup if you want to see logs.
pub fn init_logging() {
    use tracing_subscriber::{fmt, prelude::*, EnvFilter};

    tracing_subscriber::registry()
        .with(fmt::layer())
        .with(EnvFilter::from_default_env())
        .init();
}
