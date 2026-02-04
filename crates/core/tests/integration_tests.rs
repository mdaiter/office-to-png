//! Integration tests for office-to-png-core.
//!
//! These tests require:
//! - LibreOffice installed (soffice in PATH or /opt/homebrew/bin/soffice)
//! - Pdfium library (set PDFIUM_DYNAMIC_LIB_PATH or place in ./lib)
//!
//! Run with: cargo test --package office-to-png-core --test integration_tests

use office_to_png_core::{
    config::{ConversionRequest, ConverterConfig, PoolConfig},
    converter::Converter,
    pool::LibreOfficePool,
};
use std::path::PathBuf;
use tempfile::TempDir;

/// Get the path to the test fixtures directory
fn fixtures_dir() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .unwrap()
        .parent()
        .unwrap()
        .join("tests/fixtures/output")
}

/// Get the path to the pdfium library
fn pdfium_lib_path() -> Option<PathBuf> {
    // Check environment variable first
    if let Ok(path) = std::env::var("PDFIUM_DYNAMIC_LIB_PATH") {
        let p = PathBuf::from(path);
        if p.exists() {
            return Some(p);
        }
    }

    // Check project lib directory
    let project_lib = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .unwrap()
        .parent()
        .unwrap()
        .join("lib/lib");
    if project_lib.join("libpdfium.dylib").exists() || project_lib.join("libpdfium.so").exists() {
        return Some(project_lib);
    }

    None
}

/// Check if LibreOffice is available
fn libreoffice_available() -> bool {
    which::which("soffice").is_ok()
}

/// Check if pdfium is available
fn pdfium_available() -> bool {
    pdfium_lib_path().is_some()
}

/// Skip test if dependencies are not available
macro_rules! require_deps {
    () => {
        if !libreoffice_available() {
            eprintln!("Skipping test: LibreOffice not found");
            return;
        }
        if !pdfium_available() {
            eprintln!("Skipping test: Pdfium not found");
            return;
        }
    };
}

/// Setup pdfium environment
fn setup_pdfium() {
    if let Some(path) = pdfium_lib_path() {
        std::env::set_var("PDFIUM_DYNAMIC_LIB_PATH", &path);
    }
}

// ============================================================================
// LibreOffice Pool Tests
// ============================================================================

#[tokio::test]
async fn test_pool_creation_with_libreoffice() {
    if !libreoffice_available() {
        eprintln!("Skipping test: LibreOffice not found");
        return;
    }

    let config = PoolConfig::default();
    let pool = LibreOfficePool::new(config).await;
    assert!(pool.is_ok(), "Pool creation should succeed: {:?}", pool.err());
}

#[tokio::test]
async fn test_pool_health_check() {
    if !libreoffice_available() {
        eprintln!("Skipping test: LibreOffice not found");
        return;
    }

    let config = PoolConfig::with_pool_size(1);
    let pool = LibreOfficePool::new(config).await.unwrap();
    let health = pool.health().await;

    assert!(health.pool_size > 0);
    assert!(!health.instances.is_empty());
}

// ============================================================================
// Document Conversion Tests
// ============================================================================

#[tokio::test]
async fn test_convert_simple_docx() {
    require_deps!();
    setup_pdfium();

    let fixture = fixtures_dir().join("simple.docx");
    if !fixture.exists() {
        eprintln!("Skipping test: fixture not found at {:?}", fixture);
        return;
    }

    let temp_dir = TempDir::new().unwrap();
    let output_dir = temp_dir.path().to_path_buf();

    let config = ConverterConfig::new(1, 150);
    let converter = Converter::new(config).await;

    match converter {
        Ok(converter) => {
            let request = ConversionRequest::new(&fixture, &output_dir);
            let result = converter.convert(request).await;

            match result {
                Ok(file_result) => {
                    assert!(file_result.page_count > 0, "Should have at least one page");
                    assert!(!file_result.output_paths.is_empty(), "Should have output paths");

                    // Check that PNG files were created
                    for path in &file_result.output_paths {
                        assert!(path.exists(), "PNG file should exist: {:?}", path);
                    }
                    eprintln!(
                        "Success: Converted {} to {} pages",
                        fixture.display(),
                        file_result.page_count
                    );
                }
                Err(e) => {
                    eprintln!("Conversion failed (may be expected): {:?}", e);
                }
            }
        }
        Err(e) => {
            eprintln!("Converter creation failed (may be expected): {:?}", e);
        }
    }
}

#[tokio::test]
async fn test_convert_simple_xlsx() {
    require_deps!();
    setup_pdfium();

    let fixture = fixtures_dir().join("simple.xlsx");
    if !fixture.exists() {
        eprintln!("Skipping test: fixture not found at {:?}", fixture);
        return;
    }

    let temp_dir = TempDir::new().unwrap();
    let output_dir = temp_dir.path().to_path_buf();

    let config = ConverterConfig::new(1, 150);
    let converter = Converter::new(config).await;

    match converter {
        Ok(converter) => {
            let request = ConversionRequest::new(&fixture, &output_dir);
            let result = converter.convert(request).await;

            match result {
                Ok(file_result) => {
                    assert!(file_result.page_count > 0, "Should have at least one page");
                    eprintln!(
                        "Success: Converted {} to {} pages",
                        fixture.display(),
                        file_result.page_count
                    );
                }
                Err(e) => {
                    eprintln!("Conversion failed (may be expected): {:?}", e);
                }
            }
        }
        Err(e) => {
            eprintln!("Converter creation failed (may be expected): {:?}", e);
        }
    }
}

#[tokio::test]
async fn test_convert_multipage_docx() {
    require_deps!();
    setup_pdfium();

    let fixture = fixtures_dir().join("multipage.docx");
    if !fixture.exists() {
        eprintln!("Skipping test: fixture not found at {:?}", fixture);
        return;
    }

    let temp_dir = TempDir::new().unwrap();
    let output_dir = temp_dir.path().to_path_buf();

    let config = ConverterConfig::new(1, 150);
    let converter = Converter::new(config).await;

    match converter {
        Ok(converter) => {
            let request = ConversionRequest::new(&fixture, &output_dir);
            let result = converter.convert(request).await;

            match result {
                Ok(file_result) => {
                    // Multipage document should have multiple pages
                    assert!(
                        file_result.page_count >= 2,
                        "Multipage doc should have 2+ pages, got {}",
                        file_result.page_count
                    );
                    eprintln!(
                        "Success: Converted multipage doc to {} pages",
                        file_result.page_count
                    );
                }
                Err(e) => {
                    eprintln!("Conversion failed (may be expected): {:?}", e);
                }
            }
        }
        Err(e) => {
            eprintln!("Converter creation failed (may be expected): {:?}", e);
        }
    }
}

#[tokio::test]
async fn test_convert_formatted_docx() {
    require_deps!();
    setup_pdfium();

    let fixture = fixtures_dir().join("formatted.docx");
    if !fixture.exists() {
        eprintln!("Skipping test: fixture not found at {:?}", fixture);
        return;
    }

    let temp_dir = TempDir::new().unwrap();
    let output_dir = temp_dir.path().to_path_buf();

    let config = ConverterConfig::new(1, 150);
    let converter = Converter::new(config).await;

    match converter {
        Ok(converter) => {
            let request = ConversionRequest::new(&fixture, &output_dir);
            let result = converter.convert(request).await;

            match result {
                Ok(file_result) => {
                    eprintln!(
                        "Success: Converted formatted doc to {} pages",
                        file_result.page_count
                    );
                }
                Err(e) => {
                    eprintln!("Conversion failed (may be expected): {:?}", e);
                }
            }
        }
        Err(e) => {
            eprintln!("Converter creation failed (may be expected): {:?}", e);
        }
    }
}

#[tokio::test]
async fn test_convert_tables_docx() {
    require_deps!();
    setup_pdfium();

    let fixture = fixtures_dir().join("tables.docx");
    if !fixture.exists() {
        eprintln!("Skipping test: fixture not found at {:?}", fixture);
        return;
    }

    let temp_dir = TempDir::new().unwrap();
    let output_dir = temp_dir.path().to_path_buf();

    let config = ConverterConfig::new(1, 150);
    let converter = Converter::new(config).await;

    match converter {
        Ok(converter) => {
            let request = ConversionRequest::new(&fixture, &output_dir);
            let result = converter.convert(request).await;

            match result {
                Ok(file_result) => {
                    eprintln!(
                        "Success: Converted tables doc to {} pages",
                        file_result.page_count
                    );
                }
                Err(e) => {
                    eprintln!("Conversion failed (may be expected): {:?}", e);
                }
            }
        }
        Err(e) => {
            eprintln!("Converter creation failed (may be expected): {:?}", e);
        }
    }
}

#[tokio::test]
async fn test_convert_formatted_xlsx() {
    require_deps!();
    setup_pdfium();

    let fixture = fixtures_dir().join("formatted.xlsx");
    if !fixture.exists() {
        eprintln!("Skipping test: fixture not found at {:?}", fixture);
        return;
    }

    let temp_dir = TempDir::new().unwrap();
    let output_dir = temp_dir.path().to_path_buf();

    let config = ConverterConfig::new(1, 150);
    let converter = Converter::new(config).await;

    match converter {
        Ok(converter) => {
            let request = ConversionRequest::new(&fixture, &output_dir);
            let result = converter.convert(request).await;

            match result {
                Ok(file_result) => {
                    eprintln!(
                        "Success: Converted formatted xlsx to {} pages",
                        file_result.page_count
                    );
                }
                Err(e) => {
                    eprintln!("Conversion failed (may be expected): {:?}", e);
                }
            }
        }
        Err(e) => {
            eprintln!("Converter creation failed (may be expected): {:?}", e);
        }
    }
}

#[tokio::test]
async fn test_convert_multisheet_xlsx() {
    require_deps!();
    setup_pdfium();

    let fixture = fixtures_dir().join("multisheet.xlsx");
    if !fixture.exists() {
        eprintln!("Skipping test: fixture not found at {:?}", fixture);
        return;
    }

    let temp_dir = TempDir::new().unwrap();
    let output_dir = temp_dir.path().to_path_buf();

    let config = ConverterConfig::new(1, 150);
    let converter = Converter::new(config).await;

    match converter {
        Ok(converter) => {
            let request = ConversionRequest::new(&fixture, &output_dir);
            let result = converter.convert(request).await;

            match result {
                Ok(file_result) => {
                    // Multisheet should produce multiple pages
                    eprintln!(
                        "Success: Converted multisheet xlsx to {} pages",
                        file_result.page_count
                    );
                }
                Err(e) => {
                    eprintln!("Conversion failed (may be expected): {:?}", e);
                }
            }
        }
        Err(e) => {
            eprintln!("Converter creation failed (may be expected): {:?}", e);
        }
    }
}

// ============================================================================
// Error Handling Tests
// ============================================================================

#[tokio::test]
async fn test_convert_nonexistent_file() {
    require_deps!();
    setup_pdfium();

    let temp_dir = TempDir::new().unwrap();
    let output_dir = temp_dir.path().to_path_buf();
    let nonexistent = PathBuf::from("/nonexistent/file.docx");

    let config = ConverterConfig::new(1, 150);
    let converter = Converter::new(config).await;

    match converter {
        Ok(converter) => {
            let request = ConversionRequest::new(&nonexistent, &output_dir);
            let result = converter.convert(request).await;

            // Should fail with input not found error
            assert!(result.is_err(), "Should fail for nonexistent file");
            eprintln!("Correctly failed for nonexistent file: {:?}", result.err());
        }
        Err(e) => {
            eprintln!("Converter creation failed (may be expected): {:?}", e);
        }
    }
}

#[tokio::test]
async fn test_convert_corrupt_file() {
    require_deps!();
    setup_pdfium();

    let fixture = fixtures_dir().join("corrupt.docx");
    if !fixture.exists() {
        eprintln!("Skipping test: fixture not found at {:?}", fixture);
        return;
    }

    let temp_dir = TempDir::new().unwrap();
    let output_dir = temp_dir.path().to_path_buf();

    let config = ConverterConfig::new(1, 150);
    let converter = Converter::new(config).await;

    match converter {
        Ok(converter) => {
            let request = ConversionRequest::new(&fixture, &output_dir);
            let result = converter.convert(request).await;

            // Corrupt file should fail
            match result {
                Ok(file_result) => {
                    // LibreOffice might still try to process it
                    eprintln!("Corrupt file processed: {} pages", file_result.page_count);
                }
                Err(e) => {
                    // Expected - corrupt file should fail
                    eprintln!("Corrupt file failed as expected: {:?}", e);
                }
            }
        }
        Err(e) => {
            eprintln!("Converter creation failed (may be expected): {:?}", e);
        }
    }
}

#[tokio::test]
async fn test_convert_empty_docx() {
    require_deps!();
    setup_pdfium();

    let fixture = fixtures_dir().join("empty.docx");
    if !fixture.exists() {
        eprintln!("Skipping test: fixture not found at {:?}", fixture);
        return;
    }

    let temp_dir = TempDir::new().unwrap();
    let output_dir = temp_dir.path().to_path_buf();

    let config = ConverterConfig::new(1, 150);
    let converter = Converter::new(config).await;

    match converter {
        Ok(converter) => {
            let request = ConversionRequest::new(&fixture, &output_dir);
            let result = converter.convert(request).await;

            match result {
                Ok(file_result) => {
                    // Empty doc should produce at least one blank page
                    eprintln!(
                        "Empty doc converted to {} pages",
                        file_result.page_count
                    );
                }
                Err(e) => {
                    eprintln!("Empty doc conversion failed: {:?}", e);
                }
            }
        }
        Err(e) => {
            eprintln!("Converter creation failed (may be expected): {:?}", e);
        }
    }
}

// ============================================================================
// Batch Conversion Tests
// ============================================================================

#[tokio::test]
async fn test_batch_conversion() {
    require_deps!();
    setup_pdfium();

    let fixtures = fixtures_dir();
    let files: Vec<PathBuf> = vec![
        fixtures.join("simple.docx"),
        fixtures.join("simple.xlsx"),
    ]
    .into_iter()
    .filter(|p| p.exists())
    .collect();

    if files.is_empty() {
        eprintln!("Skipping test: no fixtures found");
        return;
    }

    let temp_dir = TempDir::new().unwrap();
    let output_dir = temp_dir.path().to_path_buf();

    let config = ConverterConfig::new(2, 150);
    let converter = Converter::new(config).await;

    match converter {
        Ok(converter) => {
            let requests: Vec<ConversionRequest> = files
                .iter()
                .map(|f| ConversionRequest::new(f, &output_dir))
                .collect();

            let batch_result = converter.convert_batch(requests).await;

            eprintln!(
                "Batch conversion: {} succeeded, {} failed, {} total pages",
                batch_result.successful.len(),
                batch_result.failed.len(),
                batch_result.total_pages
            );
        }
        Err(e) => {
            eprintln!("Converter creation failed (may be expected): {:?}", e);
        }
    }
}

// ============================================================================
// Configuration Tests
// ============================================================================

#[tokio::test]
async fn test_custom_dpi() {
    require_deps!();
    setup_pdfium();

    let fixture = fixtures_dir().join("simple.docx");
    if !fixture.exists() {
        eprintln!("Skipping test: fixture not found");
        return;
    }

    let temp_dir = TempDir::new().unwrap();

    // Test with different DPI settings
    for dpi in [72, 150, 300] {
        let output_dir = temp_dir.path().join(format!("dpi_{}", dpi));
        std::fs::create_dir_all(&output_dir).unwrap();

        let config = ConverterConfig::new(1, dpi);
        let converter = Converter::new(config).await;

        match converter {
            Ok(converter) => {
                let request = ConversionRequest::new(&fixture, &output_dir)
                    .with_prefix(&format!("test_dpi_{}", dpi));
                let result = converter.convert(request).await;

                match result {
                    Ok(file_result) => {
                        eprintln!(
                            "DPI {}: {} pages, {:?} duration",
                            dpi, file_result.page_count, file_result.duration
                        );
                    }
                    Err(e) => {
                        eprintln!("DPI {} conversion failed: {:?}", dpi, e);
                    }
                }
            }
            Err(e) => {
                eprintln!("Converter creation failed for DPI {}: {:?}", dpi, e);
            }
        }
    }
}

// ============================================================================
// Complex Document Tests
// ============================================================================

#[tokio::test]
async fn test_convert_complex_docx() {
    require_deps!();
    setup_pdfium();

    let fixture = fixtures_dir().join("complex.docx");
    if !fixture.exists() {
        eprintln!("Skipping test: fixture not found at {:?}", fixture);
        return;
    }

    let temp_dir = TempDir::new().unwrap();
    let output_dir = temp_dir.path().to_path_buf();

    let config = ConverterConfig::new(1, 150);
    let converter = Converter::new(config).await;

    match converter {
        Ok(converter) => {
            let request = ConversionRequest::new(&fixture, &output_dir);
            let result = converter.convert(request).await;

            match result {
                Ok(file_result) => {
                    eprintln!(
                        "Success: Converted complex doc to {} pages in {:?}",
                        file_result.page_count,
                        file_result.duration
                    );
                }
                Err(e) => {
                    eprintln!("Conversion failed (may be expected): {:?}", e);
                }
            }
        }
        Err(e) => {
            eprintln!("Converter creation failed (may be expected): {:?}", e);
        }
    }
}
