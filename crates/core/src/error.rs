//! Error types for office-to-png conversion.

use std::path::PathBuf;
use thiserror::Error;

/// Main error type for the office-to-png library.
#[derive(Error, Debug)]
pub enum ConversionError {
    /// LibreOffice is not installed or not found in PATH.
    #[error("LibreOffice not found. Please install LibreOffice and ensure 'soffice' is in PATH")]
    LibreOfficeNotFound,

    /// LibreOffice process failed to start.
    #[error("Failed to start LibreOffice process: {0}")]
    ProcessStartFailed(#[from] std::io::Error),

    /// LibreOffice conversion failed.
    #[error("LibreOffice conversion failed for '{path}': {message}")]
    ConversionFailed { path: PathBuf, message: String },

    /// LibreOffice process timed out.
    #[error("LibreOffice conversion timed out after {timeout_secs} seconds for '{path}'")]
    Timeout { path: PathBuf, timeout_secs: u64 },

    /// Input file not found.
    #[error("Input file not found: {0}")]
    InputNotFound(PathBuf),

    /// Unsupported file format.
    #[error("Unsupported file format: {extension}. Supported: .docx, .doc, .xlsx, .xls")]
    UnsupportedFormat { extension: String },

    /// PDF rendering failed.
    #[error("PDF rendering failed: {0}")]
    PdfRenderError(String),

    /// Pdfium library error.
    #[error("Pdfium error: {0}")]
    PdfiumError(String),

    /// PNG encoding failed.
    #[error("PNG encoding failed: {0}")]
    PngEncodingError(String),

    /// Output directory creation failed.
    #[error("Failed to create output directory '{path}': {message}")]
    OutputDirError { path: PathBuf, message: String },

    /// Pool exhausted - no available LibreOffice instances.
    #[error("LibreOffice pool exhausted, all {pool_size} instances are busy")]
    PoolExhausted { pool_size: usize },

    /// Pool shutdown.
    #[error("LibreOffice pool has been shut down")]
    PoolShutdown,

    /// Instance health check failed.
    #[error("LibreOffice instance health check failed: {0}")]
    HealthCheckFailed(String),

    /// Invalid configuration.
    #[error("Invalid configuration: {0}")]
    InvalidConfig(String),

    /// Channel communication error.
    #[error("Internal channel error: {0}")]
    ChannelError(String),
}

/// Result type alias for convenience.
pub type Result<T> = std::result::Result<T, ConversionError>;

impl From<async_channel::RecvError> for ConversionError {
    fn from(_: async_channel::RecvError) -> Self {
        ConversionError::ChannelError("Channel closed".to_string())
    }
}

impl<T> From<async_channel::SendError<T>> for ConversionError {
    fn from(_: async_channel::SendError<T>) -> Self {
        ConversionError::ChannelError("Channel closed".to_string())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_error_display_libreoffice_not_found() {
        let err = ConversionError::LibreOfficeNotFound;
        let msg = format!("{}", err);
        assert!(msg.contains("LibreOffice not found"));
        assert!(msg.contains("soffice"));
    }

    #[test]
    fn test_error_display_conversion_failed() {
        let err = ConversionError::ConversionFailed {
            path: PathBuf::from("/path/to/doc.docx"),
            message: "Invalid format".to_string(),
        };
        let msg = format!("{}", err);
        assert!(msg.contains("/path/to/doc.docx"));
        assert!(msg.contains("Invalid format"));
    }

    #[test]
    fn test_error_display_timeout() {
        let err = ConversionError::Timeout {
            path: PathBuf::from("doc.docx"),
            timeout_secs: 120,
        };
        let msg = format!("{}", err);
        assert!(msg.contains("120 seconds"));
        assert!(msg.contains("doc.docx"));
    }

    #[test]
    fn test_error_display_input_not_found() {
        let err = ConversionError::InputNotFound(PathBuf::from("/missing/file.docx"));
        let msg = format!("{}", err);
        assert!(msg.contains("/missing/file.docx"));
    }

    #[test]
    fn test_error_display_unsupported_format() {
        let err = ConversionError::UnsupportedFormat {
            extension: "pdf".to_string(),
        };
        let msg = format!("{}", err);
        assert!(msg.contains("pdf"));
        assert!(msg.contains("Supported"));
    }

    #[test]
    fn test_error_display_pool_exhausted() {
        let err = ConversionError::PoolExhausted { pool_size: 8 };
        let msg = format!("{}", err);
        assert!(msg.contains("8"));
        assert!(msg.contains("exhausted"));
    }

    #[test]
    fn test_error_display_invalid_config() {
        let err = ConversionError::InvalidConfig("pool_size must be > 0".to_string());
        let msg = format!("{}", err);
        assert!(msg.contains("pool_size must be > 0"));
    }

    #[test]
    fn test_error_from_io_error() {
        let io_err = std::io::Error::new(std::io::ErrorKind::NotFound, "file not found");
        let err: ConversionError = io_err.into();
        match err {
            ConversionError::ProcessStartFailed(_) => (),
            _ => panic!("Expected ProcessStartFailed"),
        }
    }

    #[test]
    fn test_error_from_recv_error() {
        let recv_err = async_channel::RecvError;
        let err: ConversionError = recv_err.into();
        match err {
            ConversionError::ChannelError(msg) => {
                assert!(msg.contains("closed"));
            }
            _ => panic!("Expected ChannelError"),
        }
    }

    #[test]
    fn test_error_debug_impl() {
        let err = ConversionError::PdfiumError("test error".to_string());
        let debug = format!("{:?}", err);
        assert!(debug.contains("PdfiumError"));
        assert!(debug.contains("test error"));
    }

    #[test]
    fn test_result_type_alias() {
        fn returns_result() -> Result<i32> {
            Ok(42)
        }
        assert_eq!(returns_result().unwrap(), 42);

        fn returns_error() -> Result<i32> {
            Err(ConversionError::PoolShutdown)
        }
        assert!(returns_error().is_err());
    }
}
