//! Test fixture generator for office-to-png.
//!
//! This binary generates test documents (DOCX and XLSX) programmatically
//! for use in unit and integration tests.

use anyhow::Result;
use docx_rs::{
    AlignmentType, BreakType, Docx, Paragraph, Run, Table as DocxTable, TableCell, TableRow,
};
use rust_xlsxwriter::{Color as XlsxColor, Format, FormatAlign, FormatBorder, Workbook};
use std::fs::{self, File};
use std::io::Write;
use std::path::Path;

fn main() -> Result<()> {
    let output_dir = Path::new("tests/fixtures/output");
    fs::create_dir_all(output_dir)?;

    println!("Generating test fixtures...\n");

    // Generate DOCX files
    generate_simple_docx(output_dir)?;
    generate_formatted_docx(output_dir)?;
    generate_multipage_docx(output_dir)?;
    generate_tables_docx(output_dir)?;
    generate_complex_docx(output_dir)?;

    // Generate XLSX files
    generate_simple_xlsx(output_dir)?;
    generate_formatted_xlsx(output_dir)?;
    generate_multisheet_xlsx(output_dir)?;

    // Generate error test files
    generate_corrupt_docx(output_dir)?;
    generate_empty_docx(output_dir)?;

    println!("\nAll fixtures generated successfully!");
    Ok(())
}

/// Generate a simple single-paragraph DOCX.
fn generate_simple_docx(output_dir: &Path) -> Result<()> {
    let path = output_dir.join("simple.docx");
    println!("  Creating: {}", path.display());

    let docx = Docx::new()
        .add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("Hello, World! This is a simple test document.")),
        )
        .add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("It contains two paragraphs of plain text.")),
        );

    let file = File::create(&path)?;
    docx.build().pack(file)?;
    Ok(())
}

/// Generate a DOCX with various text formatting.
fn generate_formatted_docx(output_dir: &Path) -> Result<()> {
    let path = output_dir.join("formatted.docx");
    println!("  Creating: {}", path.display());

    let docx = Docx::new()
        // Title
        .add_paragraph(
            Paragraph::new()
                .add_run(
                    Run::new()
                        .add_text("Formatted Document Test")
                        .bold()
                        .size(48), // 24pt (half-points)
                )
                .align(AlignmentType::Center),
        )
        // Bold text
        .add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("This text is "))
                .add_run(Run::new().add_text("bold").bold())
                .add_run(Run::new().add_text(".")),
        )
        // Italic text
        .add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("This text is "))
                .add_run(Run::new().add_text("italic").italic())
                .add_run(Run::new().add_text(".")),
        )
        // Underlined text
        .add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("This text is "))
                .add_run(Run::new().add_text("underlined").underline("single"))
                .add_run(Run::new().add_text(".")),
        )
        // Strikethrough
        .add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("This text has "))
                .add_run(Run::new().add_text("strikethrough").strike())
                .add_run(Run::new().add_text(".")),
        )
        // Colored text
        .add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("This text is "))
                .add_run(Run::new().add_text("red").color("FF0000"))
                .add_run(Run::new().add_text(", "))
                .add_run(Run::new().add_text("green").color("00FF00"))
                .add_run(Run::new().add_text(", and "))
                .add_run(Run::new().add_text("blue").color("0000FF"))
                .add_run(Run::new().add_text(".")),
        )
        // Different font sizes
        .add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("Small ").size(16)) // 8pt
                .add_run(Run::new().add_text("Medium ").size(24)) // 12pt
                .add_run(Run::new().add_text("Large ").size(36)) // 18pt
                .add_run(Run::new().add_text("Huge").size(48)), // 24pt
        )
        // Combined formatting
        .add_paragraph(
            Paragraph::new().add_run(
                Run::new()
                    .add_text("Bold, italic, underlined, and red!")
                    .bold()
                    .italic()
                    .underline("single")
                    .color("FF0000"),
            ),
        )
        // Highlighted text
        .add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("This text has "))
                .add_run(Run::new().add_text("yellow highlight").highlight("yellow"))
                .add_run(Run::new().add_text(".")),
        )
        // Right-aligned paragraph
        .add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("This paragraph is right-aligned."))
                .align(AlignmentType::Right),
        )
        // Centered paragraph
        .add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("This paragraph is centered."))
                .align(AlignmentType::Center),
        )
        // Justified paragraph
        .add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text(
                    "This is a justified paragraph. It contains enough text to demonstrate \
                     how justification works by stretching the text to fill the full width \
                     of the page margins on both sides.",
                ))
                .align(AlignmentType::Both),
        );

    let file = File::create(&path)?;
    docx.build().pack(file)?;
    Ok(())
}

/// Generate a multi-page DOCX.
fn generate_multipage_docx(output_dir: &Path) -> Result<()> {
    let path = output_dir.join("multipage.docx");
    println!("  Creating: {}", path.display());

    let mut docx = Docx::new();

    // Add title
    docx = docx.add_paragraph(
        Paragraph::new()
            .add_run(
                Run::new()
                    .add_text("Multi-Page Document Test")
                    .bold()
                    .size(48),
            )
            .align(AlignmentType::Center),
    );

    // Add enough content for multiple pages (~50 paragraphs)
    for i in 1..=50 {
        let text = format!(
            "This is paragraph {} of the multi-page test document. Lorem ipsum dolor sit amet, \
             consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore \
             magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris \
             nisi ut aliquip ex ea commodo consequat.",
            i
        );
        docx = docx.add_paragraph(Paragraph::new().add_run(Run::new().add_text(&text)));
    }

    // Add explicit page break before final section
    docx = docx.add_paragraph(Paragraph::new().add_run(Run::new().add_break(BreakType::Page)));

    docx = docx.add_paragraph(
        Paragraph::new()
            .add_run(Run::new().add_text("Final Page").bold().size(36))
            .align(AlignmentType::Center),
    );

    docx = docx.add_paragraph(
        Paragraph::new().add_run(Run::new().add_text("This is the last page of the document.")),
    );

    let file = File::create(&path)?;
    docx.build().pack(file)?;
    Ok(())
}

/// Generate a DOCX with tables.
fn generate_tables_docx(output_dir: &Path) -> Result<()> {
    let path = output_dir.join("tables.docx");
    println!("  Creating: {}", path.display());

    // Simple 3x3 table
    let simple_table = DocxTable::new(vec![
        TableRow::new(vec![
            TableCell::new()
                .add_paragraph(Paragraph::new().add_run(Run::new().add_text("A1").bold())),
            TableCell::new()
                .add_paragraph(Paragraph::new().add_run(Run::new().add_text("B1").bold())),
            TableCell::new()
                .add_paragraph(Paragraph::new().add_run(Run::new().add_text("C1").bold())),
        ]),
        TableRow::new(vec![
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("A2"))),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("B2"))),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("C2"))),
        ]),
        TableRow::new(vec![
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("A3"))),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("B3"))),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("C3"))),
        ]),
    ])
    .set_grid(vec![2000, 2000, 2000]);

    // Table with merged cells (horizontal merge)
    let merged_table = DocxTable::new(vec![
        TableRow::new(vec![TableCell::new()
            .add_paragraph(Paragraph::new().add_run(Run::new().add_text("Merged Header").bold()))
            .grid_span(3)]),
        TableRow::new(vec![
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("Col 1"))),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("Col 2"))),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("Col 3"))),
        ]),
    ])
    .set_grid(vec![2000, 2000, 2000]);

    let docx = Docx::new()
        .add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("Table Test Document").bold().size(48))
                .align(AlignmentType::Center),
        )
        .add_paragraph(Paragraph::new().add_run(Run::new().add_text("Simple 3x3 Table:").bold()))
        .add_table(simple_table)
        .add_paragraph(Paragraph::new()) // Spacing
        .add_paragraph(
            Paragraph::new().add_run(Run::new().add_text("Table with Merged Header:").bold()),
        )
        .add_table(merged_table);

    let file = File::create(&path)?;
    docx.build().pack(file)?;
    Ok(())
}

/// Generate a complex DOCX with multiple features.
fn generate_complex_docx(output_dir: &Path) -> Result<()> {
    let path = output_dir.join("complex.docx");
    println!("  Creating: {}", path.display());

    let table = DocxTable::new(vec![
        TableRow::new(vec![
            TableCell::new()
                .add_paragraph(Paragraph::new().add_run(Run::new().add_text("Item").bold())),
            TableCell::new()
                .add_paragraph(Paragraph::new().add_run(Run::new().add_text("Quantity").bold())),
            TableCell::new()
                .add_paragraph(Paragraph::new().add_run(Run::new().add_text("Price").bold())),
        ]),
        TableRow::new(vec![
            TableCell::new()
                .add_paragraph(Paragraph::new().add_run(Run::new().add_text("Widget A"))),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("10"))),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("$5.00"))),
        ]),
        TableRow::new(vec![
            TableCell::new()
                .add_paragraph(Paragraph::new().add_run(Run::new().add_text("Widget B"))),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("25"))),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("$3.50"))),
        ]),
    ])
    .set_grid(vec![3000, 1500, 1500]);

    let docx = Docx::new()
        // Title
        .add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("Complex Document Test").bold().size(56))
                .align(AlignmentType::Center),
        )
        // Subtitle
        .add_paragraph(
            Paragraph::new()
                .add_run(
                    Run::new()
                        .add_text("A comprehensive test of document features")
                        .italic()
                        .size(28),
                )
                .align(AlignmentType::Center),
        )
        .add_paragraph(Paragraph::new()) // Spacing
        // Introduction
        .add_paragraph(
            Paragraph::new().add_run(Run::new().add_text("1. Introduction").bold().size(32)),
        )
        .add_paragraph(Paragraph::new().add_run(Run::new().add_text(
            "This document tests various formatting features including text styles, \
                 colors, tables, and paragraph formatting. It serves as a comprehensive \
                 test fixture for the office-to-png conversion library.",
        )))
        // Formatted section
        .add_paragraph(
            Paragraph::new().add_run(Run::new().add_text("2. Text Formatting").bold().size(32)),
        )
        .add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("Key features: "))
                .add_run(Run::new().add_text("bold").bold())
                .add_run(Run::new().add_text(", "))
                .add_run(Run::new().add_text("italic").italic())
                .add_run(Run::new().add_text(", "))
                .add_run(Run::new().add_text("underline").underline("single"))
                .add_run(Run::new().add_text(", "))
                .add_run(Run::new().add_text("colors").color("0066CC"))
                .add_run(Run::new().add_text(", and "))
                .add_run(Run::new().add_text("highlighting").highlight("yellow"))
                .add_run(Run::new().add_text(".")),
        )
        // Table section
        .add_paragraph(Paragraph::new().add_run(Run::new().add_text("3. Tables").bold().size(32)))
        .add_paragraph(Paragraph::new().add_run(Run::new().add_text("Sample data table:")))
        .add_table(table)
        // Conclusion
        .add_paragraph(Paragraph::new())
        .add_paragraph(
            Paragraph::new().add_run(Run::new().add_text("4. Conclusion").bold().size(32)),
        )
        .add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text(
                    "This document demonstrates that the conversion library can handle \
                     a variety of common document elements.",
                ))
                .align(AlignmentType::Both),
        );

    let file = File::create(&path)?;
    docx.build().pack(file)?;
    Ok(())
}

/// Generate a simple XLSX with basic data.
fn generate_simple_xlsx(output_dir: &Path) -> Result<()> {
    let path = output_dir.join("simple.xlsx");
    println!("  Creating: {}", path.display());

    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Header row
    worksheet.write_string(0, 0, "Name")?;
    worksheet.write_string(0, 1, "Age")?;
    worksheet.write_string(0, 2, "City")?;

    // Data rows
    let data = [
        ("Alice", 28, "New York"),
        ("Bob", 35, "Los Angeles"),
        ("Charlie", 42, "Chicago"),
        ("Diana", 31, "Houston"),
        ("Eve", 26, "Phoenix"),
    ];

    for (row, (name, age, city)) in data.iter().enumerate() {
        let r = (row + 1) as u32;
        worksheet.write_string(r, 0, *name)?;
        worksheet.write_number(r, 1, *age as f64)?;
        worksheet.write_string(r, 2, *city)?;
    }

    workbook.save(&path)?;
    Ok(())
}

/// Generate an XLSX with cell formatting.
fn generate_formatted_xlsx(output_dir: &Path) -> Result<()> {
    let path = output_dir.join("formatted.xlsx");
    println!("  Creating: {}", path.display());

    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Define formats
    let bold_format = Format::new().set_bold();
    let red_format = Format::new().set_font_color(XlsxColor::Red);
    let blue_bg_format = Format::new().set_background_color(XlsxColor::RGB(0xCCE5FF));
    let border_format = Format::new()
        .set_border(FormatBorder::Thin)
        .set_border_color(XlsxColor::Black);
    let header_format = Format::new()
        .set_bold()
        .set_background_color(XlsxColor::RGB(0x4472C4))
        .set_font_color(XlsxColor::White)
        .set_border(FormatBorder::Thin);
    let number_format = Format::new().set_num_format("#,##0.00");
    let percent_format = Format::new().set_num_format("0.0%");
    let center_format = Format::new().set_align(FormatAlign::Center);
    let large_font = Format::new().set_font_size(18.0);

    // Title
    worksheet.write_string_with_format(0, 0, "Formatted Spreadsheet Test", &large_font)?;

    // Headers with formatting
    worksheet.write_string_with_format(2, 0, "Product", &header_format)?;
    worksheet.write_string_with_format(2, 1, "Sales", &header_format)?;
    worksheet.write_string_with_format(2, 2, "Growth", &header_format)?;

    // Data with various formats
    worksheet.write_string_with_format(3, 0, "Product A", &border_format)?;
    worksheet.write_number_with_format(3, 1, 1234.56, &number_format)?;
    worksheet.write_number_with_format(3, 2, 0.125, &percent_format)?;

    worksheet.write_string_with_format(4, 0, "Product B", &border_format)?;
    worksheet.write_number_with_format(4, 1, 5678.90, &number_format)?;
    worksheet.write_number_with_format(4, 2, -0.05, &percent_format)?;

    // Colored cells
    worksheet.write_string_with_format(6, 0, "Red text", &red_format)?;
    worksheet.write_string_with_format(6, 1, "Blue background", &blue_bg_format)?;
    worksheet.write_string_with_format(6, 2, "Bold text", &bold_format)?;

    // Centered text
    worksheet.write_string_with_format(8, 0, "Centered", &center_format)?;
    worksheet.write_string_with_format(8, 1, "Centered", &center_format)?;
    worksheet.write_string_with_format(8, 2, "Centered", &center_format)?;

    // Set column widths
    worksheet.set_column_width(0, 15.0)?;
    worksheet.set_column_width(1, 12.0)?;
    worksheet.set_column_width(2, 12.0)?;

    workbook.save(&path)?;
    Ok(())
}

/// Generate an XLSX with multiple sheets.
fn generate_multisheet_xlsx(output_dir: &Path) -> Result<()> {
    let path = output_dir.join("multisheet.xlsx");
    println!("  Creating: {}", path.display());

    let mut workbook = Workbook::new();

    // Sheet 1: Sales Data
    let sheet1 = workbook.add_worksheet().set_name("Sales")?;
    sheet1.write_string(0, 0, "Month")?;
    sheet1.write_string(0, 1, "Revenue")?;
    let months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"];
    let revenues = [10000.0, 12000.0, 11500.0, 13000.0, 14500.0, 15000.0];
    for (i, (month, revenue)) in months.iter().zip(revenues.iter()).enumerate() {
        let row = (i + 1) as u32;
        sheet1.write_string(row, 0, *month)?;
        sheet1.write_number(row, 1, *revenue)?;
    }

    // Sheet 2: Expenses
    let sheet2 = workbook.add_worksheet().set_name("Expenses")?;
    sheet2.write_string(0, 0, "Category")?;
    sheet2.write_string(0, 1, "Amount")?;
    let categories = ["Rent", "Utilities", "Salaries", "Marketing", "Other"];
    let amounts = [5000.0, 800.0, 25000.0, 3000.0, 1500.0];
    for (i, (cat, amt)) in categories.iter().zip(amounts.iter()).enumerate() {
        let row = (i + 1) as u32;
        sheet2.write_string(row, 0, *cat)?;
        sheet2.write_number(row, 1, *amt)?;
    }

    // Sheet 3: Summary
    let sheet3 = workbook.add_worksheet().set_name("Summary")?;
    let bold = Format::new().set_bold();
    sheet3.write_string_with_format(0, 0, "Summary Report", &bold)?;
    sheet3.write_string(2, 0, "Total Revenue")?;
    let total_revenue: f64 = revenues.iter().sum();
    sheet3.write_number(2, 1, total_revenue)?;
    sheet3.write_string(3, 0, "Total Expenses")?;
    let total_expenses: f64 = amounts.iter().sum();
    sheet3.write_number(3, 1, total_expenses)?;

    workbook.save(&path)?;
    Ok(())
}

/// Generate a corrupt DOCX for error handling tests.
fn generate_corrupt_docx(output_dir: &Path) -> Result<()> {
    let path = output_dir.join("corrupt.docx");
    println!("  Creating: {}", path.display());

    // Write invalid data (not a valid ZIP/DOCX)
    let mut file = File::create(&path)?;
    file.write_all(b"This is not a valid DOCX file. It's just garbage data.")?;
    Ok(())
}

/// Generate an empty DOCX.
fn generate_empty_docx(output_dir: &Path) -> Result<()> {
    let path = output_dir.join("empty.docx");
    println!("  Creating: {}", path.display());

    let docx = Docx::new();
    let file = File::create(&path)?;
    docx.build().pack(file)?;
    Ok(())
}
