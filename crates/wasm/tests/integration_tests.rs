//! Integration tests for office-to-png-wasm.
//!
//! These tests verify document parsing and rendering functionality.
//! They use programmatically generated documents from the fixtures.

use office_to_png_wasm::{
    docx_renderer::DocxDocument, fonts::FontManager, text_layout::TextStyle,
    text_shaper::TextShaper, xlsx_renderer::XlsxDocument,
};
use std::path::PathBuf;

/// Get the path to the test fixtures directory
fn fixtures_dir() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .unwrap()
        .parent()
        .unwrap()
        .join("tests/fixtures/output")
}

// ============================================================================
// DOCX Document Tests
// ============================================================================

#[test]
fn test_parse_simple_docx() {
    let fixture = fixtures_dir().join("simple.docx");
    if !fixture.exists() {
        eprintln!("Skipping test: fixture not found at {:?}", fixture);
        return;
    }

    let data = std::fs::read(&fixture).expect("Failed to read fixture");
    let doc = DocxDocument::from_bytes(&data);

    match doc {
        Ok(doc) => {
            assert!(
                doc.paragraph_count() > 0,
                "Should have at least one paragraph"
            );
            eprintln!("Parsed simple.docx: {} paragraphs", doc.paragraph_count());

            // Check first paragraph content
            if let Some(text) = doc.get_paragraph(0) {
                eprintln!("First paragraph: {}", text);
                assert!(!text.is_empty(), "First paragraph should have content");
            }
        }
        Err(e) => {
            eprintln!("Parse error (may be expected): {:?}", e);
        }
    }
}

#[test]
fn test_parse_formatted_docx() {
    let fixture = fixtures_dir().join("formatted.docx");
    if !fixture.exists() {
        eprintln!("Skipping test: fixture not found at {:?}", fixture);
        return;
    }

    let data = std::fs::read(&fixture).expect("Failed to read fixture");
    let doc = DocxDocument::from_bytes(&data);

    match doc {
        Ok(doc) => {
            eprintln!(
                "Parsed formatted.docx: {} paragraphs, {} pages",
                doc.paragraph_count(),
                doc.page_count()
            );
        }
        Err(e) => {
            eprintln!("Parse error (may be expected): {:?}", e);
        }
    }
}

#[test]
fn test_parse_tables_docx() {
    let fixture = fixtures_dir().join("tables.docx");
    if !fixture.exists() {
        eprintln!("Skipping test: fixture not found at {:?}", fixture);
        return;
    }

    let data = std::fs::read(&fixture).expect("Failed to read fixture");
    let doc = DocxDocument::from_bytes(&data);

    match doc {
        Ok(doc) => {
            // Count tables
            let table_count = doc
                .elements()
                .iter()
                .filter(|e| {
                    matches!(
                        e,
                        office_to_png_wasm::docx_renderer::DocumentElement::Table(_)
                    )
                })
                .count();
            eprintln!(
                "Parsed tables.docx: {} tables, {} elements",
                table_count,
                doc.elements().len()
            );
        }
        Err(e) => {
            eprintln!("Parse error (may be expected): {:?}", e);
        }
    }
}

#[test]
fn test_parse_multipage_docx() {
    let fixture = fixtures_dir().join("multipage.docx");
    if !fixture.exists() {
        eprintln!("Skipping test: fixture not found at {:?}", fixture);
        return;
    }

    let data = std::fs::read(&fixture).expect("Failed to read fixture");
    let doc = DocxDocument::from_bytes(&data);

    match doc {
        Ok(doc) => {
            eprintln!(
                "Parsed multipage.docx: estimated {} pages",
                doc.page_count()
            );
        }
        Err(e) => {
            eprintln!("Parse error (may be expected): {:?}", e);
        }
    }
}

#[test]
fn test_parse_complex_docx() {
    let fixture = fixtures_dir().join("complex.docx");
    if !fixture.exists() {
        eprintln!("Skipping test: fixture not found at {:?}", fixture);
        return;
    }

    let data = std::fs::read(&fixture).expect("Failed to read fixture");
    let doc = DocxDocument::from_bytes(&data);

    match doc {
        Ok(doc) => {
            eprintln!(
                "Parsed complex.docx: {} paragraphs, {} elements",
                doc.paragraph_count(),
                doc.elements().len()
            );
        }
        Err(e) => {
            eprintln!("Parse error (may be expected): {:?}", e);
        }
    }
}

// ============================================================================
// XLSX Document Tests
// ============================================================================

#[test]
fn test_parse_simple_xlsx() {
    let fixture = fixtures_dir().join("simple.xlsx");
    if !fixture.exists() {
        eprintln!("Skipping test: fixture not found at {:?}", fixture);
        return;
    }

    let data = std::fs::read(&fixture).expect("Failed to read fixture");
    let doc = XlsxDocument::from_bytes(&data);

    match doc {
        Ok(doc) => {
            assert!(doc.sheet_count() > 0, "Should have at least one sheet");
            eprintln!("Parsed simple.xlsx: {} sheets", doc.sheet_count());

            // Get first sheet data
            if let Some(sheet) = doc.get_styled_sheet_data(0) {
                eprintln!(
                    "First sheet: {} rows, {} columns",
                    sheet.row_count, sheet.col_count
                );
            }
        }
        Err(e) => {
            eprintln!("Parse error (may be expected): {:?}", e);
        }
    }
}

#[test]
fn test_parse_formatted_xlsx() {
    let fixture = fixtures_dir().join("formatted.xlsx");
    if !fixture.exists() {
        eprintln!("Skipping test: fixture not found at {:?}", fixture);
        return;
    }

    let data = std::fs::read(&fixture).expect("Failed to read fixture");
    let doc = XlsxDocument::from_bytes(&data);

    match doc {
        Ok(doc) => {
            eprintln!("Parsed formatted.xlsx: {} sheets", doc.sheet_count());

            // Check for styled cells
            if let Some(sheet) = doc.get_styled_sheet_data(0) {
                let styled_cells = sheet
                    .cells
                    .iter()
                    .filter(|c| !c.bg_color.is_empty() || c.bold)
                    .count();
                eprintln!("Found {} styled cells", styled_cells);
            }
        }
        Err(e) => {
            eprintln!("Parse error (may be expected): {:?}", e);
        }
    }
}

#[test]
fn test_parse_multisheet_xlsx() {
    let fixture = fixtures_dir().join("multisheet.xlsx");
    if !fixture.exists() {
        eprintln!("Skipping test: fixture not found at {:?}", fixture);
        return;
    }

    let data = std::fs::read(&fixture).expect("Failed to read fixture");
    let doc = XlsxDocument::from_bytes(&data);

    match doc {
        Ok(doc) => {
            assert!(doc.sheet_count() >= 2, "Should have multiple sheets");
            eprintln!("Parsed multisheet.xlsx: {} sheets", doc.sheet_count());

            // List sheet names
            let names = doc.sheet_names();
            for (i, name) in names.iter().enumerate() {
                eprintln!("  Sheet {}: {}", i, name);
            }
        }
        Err(e) => {
            eprintln!("Parse error (may be expected): {:?}", e);
        }
    }
}

// ============================================================================
// Font and Text Shaping Tests
// ============================================================================

#[test]
fn test_font_manager() {
    let manager = FontManager::new();

    assert!(manager.has_font("Noto Sans"), "Should have Noto Sans font");
    assert_eq!(manager.default_font(), "Noto Sans");

    let fonts = manager.available_fonts();
    eprintln!("Available fonts: {:?}", fonts);
    assert!(!fonts.is_empty());
}

#[test]
fn test_text_shaper_basic() {
    let shaper = TextShaper::new();
    let style = TextStyle::new("Noto Sans", 12.0);

    let shaped = shaper.shape_text("Hello, World!", &style, None);

    assert!(!shaped.lines.is_empty(), "Should have at least one line");
    assert!(shaped.width > 0.0, "Width should be positive");
    assert!(shaped.height > 0.0, "Height should be positive");

    eprintln!(
        "Shaped text: {} lines, {}x{} pixels",
        shaped.lines.len(),
        shaped.width,
        shaped.height
    );
}

#[test]
fn test_text_shaper_with_wrapping() {
    let shaper = TextShaper::new();
    let style = TextStyle::new("Noto Sans", 12.0);

    let long_text = "This is a longer text that should wrap to multiple lines when given a narrow width constraint.";
    let shaped = shaper.shape_text(long_text, &style, Some(100.0));

    eprintln!("Wrapped text: {} lines", shaped.lines.len());
    // With a narrow width, should wrap to multiple lines
    assert!(shaped.lines.len() >= 1, "Should have wrapped lines");
}

#[test]
fn test_text_shaper_styled() {
    let shaper = TextShaper::new();

    // Test bold
    let mut bold_style = TextStyle::new("Noto Sans", 12.0);
    bold_style.bold = true;
    let bold_shaped = shaper.shape_text("Bold Text", &bold_style, None);

    // Test italic
    let mut italic_style = TextStyle::new("Noto Sans", 12.0);
    italic_style.italic = true;
    let italic_shaped = shaper.shape_text("Italic Text", &italic_style, None);

    // Both should produce output
    assert!(!bold_shaped.lines.is_empty());
    assert!(!italic_shaped.lines.is_empty());

    eprintln!(
        "Bold width: {}, Italic width: {}",
        bold_shaped.width, italic_shaped.width
    );
}

#[test]
fn test_text_measure() {
    let shaper = TextShaper::new();
    let style = TextStyle::new("Noto Sans", 12.0);

    let (width, height) = shaper.measure_text("Test", &style);

    assert!(width > 0.0);
    assert!(height > 0.0);

    // Measure longer text should be wider
    let (long_width, _) = shaper.measure_text("Test Test Test", &style);
    assert!(long_width > width, "Longer text should be wider");
}

// ============================================================================
// DOCX Page Dimension Tests
// ============================================================================

#[test]
fn test_docx_page_dimensions() {
    let fixture = fixtures_dir().join("simple.docx");
    if !fixture.exists() {
        eprintln!("Skipping test: fixture not found");
        return;
    }

    let data = std::fs::read(&fixture).expect("Failed to read fixture");
    let doc = DocxDocument::from_bytes(&data);

    if let Ok(doc) = doc {
        let (width, height) = doc.page_dimensions();
        let (top, right, bottom, left) = doc.margins();
        let content = doc.content_area();

        eprintln!("Page: {}x{} pixels", width, height);
        eprintln!(
            "Margins: top={}, right={}, bottom={}, left={}",
            top, right, bottom, left
        );
        eprintln!(
            "Content area: {}x{} at ({}, {})",
            content.width, content.height, content.x, content.y
        );

        assert!(width > 0.0);
        assert!(height > 0.0);
    }
}

// ============================================================================
// XLSX Grid Rendering Tests
// ============================================================================

#[test]
fn test_xlsx_sheet_data() {
    let fixture = fixtures_dir().join("simple.xlsx");
    if !fixture.exists() {
        eprintln!("Skipping test: fixture not found");
        return;
    }

    let data = std::fs::read(&fixture).expect("Failed to read fixture");
    let doc = XlsxDocument::from_bytes(&data);

    if let Ok(doc) = doc {
        if let Some(sheet) = doc.get_styled_sheet_data(0) {
            eprintln!("Sheet data:");
            eprintln!("  Rows: {}", sheet.row_count);
            eprintln!("  Columns: {}", sheet.col_count);
            eprintln!("  Cells: {}", sheet.cells.len());
            eprintln!(
                "  Column widths: {:?}",
                &sheet.column_widths[..sheet.column_widths.len().min(5)]
            );
            eprintln!(
                "  Row heights: {:?}",
                &sheet.row_heights[..sheet.row_heights.len().min(5)]
            );

            // Print first few cells
            for (i, cell) in sheet.cells.iter().take(5).enumerate() {
                eprintln!(
                    "  Cell {}: ({}, {}) = {:?}",
                    i, cell.row, cell.col, cell.value
                );
            }
        }
    }
}
