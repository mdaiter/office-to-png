//! XLSX spreadsheet parsing with styling support.
//!
//! Uses umya-spreadsheet for Excel file parsing including:
//! - Cell values
//! - Text formatting (bold, italic, colors, fonts)
//! - Cell backgrounds and borders
//! - Column widths and row heights
//! - Merged cells

use crate::renderer::{points_to_pixels, Color};
use crate::styles::{
    BorderStyle, BorderType, CellStyle, FillStyle, HorizontalAlign, VerticalAlign,
};
use crate::text_layout::TextStyle;
use serde::{Deserialize, Serialize};
use std::io::Cursor;
use umya_spreadsheet::Spreadsheet;

/// Error type for XLSX operations.
#[derive(Debug)]
pub enum XlsxError {
    ParseError(String),
}

impl std::fmt::Display for XlsxError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            XlsxError::ParseError(msg) => write!(f, "Parse error: {}", msg),
        }
    }
}

impl std::error::Error for XlsxError {}

/// Cell data with styling information.
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct StyledCell {
    /// The cell value as a string
    pub value: String,
    /// Row index (0-based)
    pub row: u32,
    /// Column index (0-based)
    pub col: u32,
    /// Font family
    pub font_family: String,
    /// Font size in points
    pub font_size: f32,
    /// Is bold
    pub bold: bool,
    /// Is italic
    pub italic: bool,
    /// Is underlined
    pub underline: bool,
    /// Text color (hex)
    pub text_color: String,
    /// Background color (hex, empty for none)
    pub bg_color: String,
    /// Horizontal alignment
    pub h_align: String,
    /// Vertical alignment
    pub v_align: String,
    /// Text wrap enabled
    pub wrap_text: bool,
    /// Number format string
    pub number_format: String,
    /// Border top (width, color, style)
    pub border_top: Option<BorderInfo>,
    /// Border right
    pub border_right: Option<BorderInfo>,
    /// Border bottom
    pub border_bottom: Option<BorderInfo>,
    /// Border left
    pub border_left: Option<BorderInfo>,
}

/// Border information
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct BorderInfo {
    pub width: f32,
    pub color: String,
    pub style: String,
}

/// Sheet data with styling
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct SheetData {
    /// Sheet name
    pub name: String,
    /// Cells with styling
    pub cells: Vec<StyledCell>,
    /// Column widths in pixels
    pub column_widths: Vec<f32>,
    /// Row heights in pixels
    pub row_heights: Vec<f32>,
    /// Merged cell ranges [(start_row, start_col, end_row, end_col), ...]
    pub merged_cells: Vec<(u32, u32, u32, u32)>,
    /// Total number of rows
    pub row_count: u32,
    /// Total number of columns
    pub col_count: u32,
}

/// A parsed XLSX spreadsheet with full styling.
pub struct XlsxDocument {
    spreadsheet: Spreadsheet,
    sheet_names: Vec<String>,
    title: Option<String>,
}

impl XlsxDocument {
    /// Parse an XLSX spreadsheet from bytes.
    pub fn from_bytes(data: &[u8]) -> Result<Self, XlsxError> {
        let cursor = Cursor::new(data);
        let spreadsheet = umya_spreadsheet::reader::xlsx::read_reader(cursor, true)
            .map_err(|e| XlsxError::ParseError(format!("Failed to parse XLSX: {}", e)))?;

        let sheet_names: Vec<String> = spreadsheet
            .get_sheet_collection()
            .iter()
            .map(|sheet| sheet.get_name().to_string())
            .collect();

        Ok(Self {
            spreadsheet,
            sheet_names,
            title: None,
        })
    }

    /// Get the number of sheets.
    pub fn sheet_count(&self) -> usize {
        self.sheet_names.len()
    }

    /// Get sheet names.
    pub fn sheet_names(&self) -> &[String] {
        &self.sheet_names
    }

    /// Get the document title.
    pub fn title(&self) -> Option<&str> {
        self.title.as_deref()
    }

    /// Get cell data for a sheet (legacy format, just values).
    pub fn get_cell_data(&self, sheet_index: usize) -> Option<Vec<Vec<String>>> {
        let sheet_name = self.sheet_names.get(sheet_index)?;
        let sheet = self.spreadsheet.get_sheet_by_name(sheet_name)?;

        let max_col = sheet.get_highest_column();
        let max_row = sheet.get_highest_row();

        let mut rows = Vec::new();

        for row_idx in 1..=max_row {
            let mut cells = Vec::new();
            for col_idx in 1..=max_col {
                let value = sheet
                    .get_cell((col_idx, row_idx))
                    .map(|cell| cell.get_value().to_string())
                    .unwrap_or_default();
                cells.push(value);
            }
            rows.push(cells);
        }

        Some(rows)
    }

    /// Get styled sheet data for rendering.
    pub fn get_styled_sheet_data(&self, sheet_index: usize) -> Option<SheetData> {
        let sheet_name = self.sheet_names.get(sheet_index)?;
        let sheet = self.spreadsheet.get_sheet_by_name(sheet_name)?;

        let max_col = sheet.get_highest_column();
        let max_row = sheet.get_highest_row();

        // Default dimensions
        const DEFAULT_COLUMN_WIDTH: f32 = 64.0; // ~8.43 characters
        const DEFAULT_ROW_HEIGHT: f32 = 20.0; // ~15 points

        // Collect column widths
        let mut column_widths = Vec::with_capacity(max_col as usize);
        for col_idx in 1..=max_col {
            let width = sheet
                .get_column_dimension_by_number(&col_idx)
                .map(|dim| {
                    let w = dim.get_width();
                    if *w > 0.0 {
                        // Excel width is in characters, roughly 7 pixels per character
                        *w as f32 * 7.0
                    } else {
                        DEFAULT_COLUMN_WIDTH
                    }
                })
                .unwrap_or(DEFAULT_COLUMN_WIDTH);
            column_widths.push(width);
        }

        // Collect row heights
        let mut row_heights = Vec::with_capacity(max_row as usize);
        for row_idx in 1..=max_row {
            let height = sheet
                .get_row_dimension(&row_idx)
                .map(|dim| {
                    let h = dim.get_height();
                    if *h > 0.0 {
                        points_to_pixels(*h as f32)
                    } else {
                        DEFAULT_ROW_HEIGHT
                    }
                })
                .unwrap_or(DEFAULT_ROW_HEIGHT);
            row_heights.push(height);
        }

        // Collect cells with styling
        let mut cells = Vec::new();

        for row_idx in 1..=max_row {
            for col_idx in 1..=max_col {
                if let Some(cell) = sheet.get_cell((col_idx, row_idx)) {
                    let value = cell.get_value().to_string();

                    // Get style from cell
                    let style = cell.get_style();

                    // Extract font info with defaults
                    let font = style.get_font();
                    let (font_family, font_size, bold, italic, underline, text_color) =
                        if let Some(f) = font {
                            let name = f.get_name().to_string();
                            let size = *f.get_size() as f32;
                            let is_bold = *f.get_bold();
                            let is_italic = *f.get_italic();
                            let is_underline = f.get_underline() != "none";
                            let argb = f.get_color().get_argb();
                            let color = if argb.is_empty() {
                                "#000000".to_string()
                            } else {
                                format!("#{}", argb)
                            };
                            (name, size, is_bold, is_italic, is_underline, color)
                        } else {
                            (
                                "Calibri".to_string(),
                                11.0,
                                false,
                                false,
                                false,
                                "#000000".to_string(),
                            )
                        };

                    // Background color
                    let fill = style.get_fill();
                    let bg_color = fill
                        .and_then(|f| f.get_pattern_fill())
                        .and_then(|pf| pf.get_foreground_color())
                        .map(|c| {
                            let argb = c.get_argb();
                            if argb.is_empty() {
                                String::new()
                            } else {
                                format!("#{}", argb)
                            }
                        })
                        .unwrap_or_default();

                    // Alignment
                    let alignment = style.get_alignment();
                    let (h_align, v_align, wrap_text) = if let Some(a) = alignment {
                        (
                            format!("{:?}", a.get_horizontal()).to_lowercase(),
                            format!("{:?}", a.get_vertical()).to_lowercase(),
                            *a.get_wrap_text(),
                        )
                    } else {
                        ("general".to_string(), "center".to_string(), false)
                    };

                    // Number format
                    let number_format = style
                        .get_number_format()
                        .map(|nf| nf.get_format_code().to_string())
                        .unwrap_or_default();

                    // Borders
                    let borders = style.get_borders();
                    let (border_top, border_right, border_bottom, border_left) =
                        if let Some(b) = borders {
                            (
                                extract_border(Some(b.get_top())),
                                extract_border(Some(b.get_right())),
                                extract_border(Some(b.get_bottom())),
                                extract_border(Some(b.get_left())),
                            )
                        } else {
                            (None, None, None, None)
                        };

                    cells.push(StyledCell {
                        value,
                        row: row_idx - 1,
                        col: col_idx - 1,
                        font_family,
                        font_size,
                        bold,
                        italic,
                        underline,
                        text_color,
                        bg_color,
                        h_align,
                        v_align,
                        wrap_text,
                        number_format,
                        border_top,
                        border_right,
                        border_bottom,
                        border_left,
                    });
                }
            }
        }

        // Collect merged cells
        let merged_cells: Vec<(u32, u32, u32, u32)> = sheet
            .get_merge_cells()
            .iter()
            .filter_map(|mc| {
                // mc.get_range() returns a String directly
                let range_str = mc.get_range();
                // Parse range like "A1:B2"
                parse_cell_range(&range_str)
            })
            .collect();

        Some(SheetData {
            name: sheet_name.clone(),
            cells,
            column_widths,
            row_heights,
            merged_cells,
            row_count: max_row,
            col_count: max_col,
        })
    }

    /// Convert styled cell to CellStyle for rendering
    pub fn cell_to_render_style(cell: &StyledCell) -> CellStyle {
        let mut text_style = TextStyle::new(&cell.font_family, cell.font_size);
        text_style.bold = cell.bold;
        text_style.italic = cell.italic;
        text_style.underline = cell.underline;

        if let Some(color) = Color::from_hex(&cell.text_color) {
            text_style.color = color.to_rgba_array();
        }

        let fill = if !cell.bg_color.is_empty() {
            Color::from_hex(&cell.bg_color)
                .map(|c| FillStyle::Solid(c.to_rgba_array()))
                .unwrap_or(FillStyle::None)
        } else {
            FillStyle::None
        };

        let h_align = match cell.h_align.as_str() {
            "left" => HorizontalAlign::Left,
            "center" => HorizontalAlign::Center,
            "right" => HorizontalAlign::Right,
            "justify" => HorizontalAlign::Justify,
            "fill" => HorizontalAlign::Fill,
            _ => HorizontalAlign::General,
        };

        let v_align = match cell.v_align.as_str() {
            "top" => VerticalAlign::Top,
            "center" => VerticalAlign::Center,
            "bottom" => VerticalAlign::Bottom,
            _ => VerticalAlign::Center,
        };

        CellStyle {
            text: text_style,
            fill,
            border_top: cell.border_top.as_ref().map(border_info_to_style),
            border_right: cell.border_right.as_ref().map(border_info_to_style),
            border_bottom: cell.border_bottom.as_ref().map(border_info_to_style),
            border_left: cell.border_left.as_ref().map(border_info_to_style),
            h_align,
            v_align,
            wrap_text: cell.wrap_text,
            number_format: Some(cell.number_format.clone()),
        }
    }
}

/// Extract border info from umya border
fn extract_border(border: Option<&umya_spreadsheet::structs::Border>) -> Option<BorderInfo> {
    let border = border?;
    let style = border.get_border_style();
    if style == "none" || style.is_empty() {
        return None;
    }

    let argb = border.get_color().get_argb();
    let color = if argb.is_empty() {
        "#000000".to_string()
    } else {
        format!("#{}", argb)
    };

    let width = match style {
        "thin" => 1.0,
        "medium" => 2.0,
        "thick" => 3.0,
        "double" => 3.0,
        "hair" => 0.5,
        _ => 1.0,
    };

    Some(BorderInfo {
        width,
        color,
        style: style.to_string(),
    })
}

/// Convert BorderInfo to rendering BorderStyle
fn border_info_to_style(info: &BorderInfo) -> BorderStyle {
    let color = Color::from_hex(&info.color)
        .map(|c| c.to_rgba_array())
        .unwrap_or([0.0, 0.0, 0.0, 1.0]);

    let border_type = match info.style.as_str() {
        "dashed" => BorderType::Dashed,
        "dotted" => BorderType::Dotted,
        "double" => BorderType::Double,
        _ => BorderType::Solid,
    };

    BorderStyle {
        width: info.width,
        color,
        style: border_type,
    }
}

/// Parse Excel cell range like "A1:B2" to (start_row, start_col, end_row, end_col)
fn parse_cell_range(range: &str) -> Option<(u32, u32, u32, u32)> {
    let parts: Vec<&str> = range.split(':').collect();
    if parts.len() != 2 {
        return None;
    }

    let (start_col, start_row) = parse_cell_ref(parts[0])?;
    let (end_col, end_row) = parse_cell_ref(parts[1])?;

    Some((start_row, start_col, end_row, end_col))
}

/// Parse Excel cell reference like "A1" to (col, row) 0-indexed
fn parse_cell_ref(cell_ref: &str) -> Option<(u32, u32)> {
    let cell_ref = cell_ref.trim();
    let mut col = 0u32;
    let mut row_str = String::new();

    for c in cell_ref.chars() {
        if c.is_ascii_alphabetic() {
            col = col * 26 + (c.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
        } else if c.is_ascii_digit() {
            row_str.push(c);
        }
    }

    let row: u32 = row_str.parse().ok()?;

    // Convert to 0-indexed
    Some((col.saturating_sub(1), row.saturating_sub(1)))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_cell_ref() {
        assert_eq!(parse_cell_ref("A1"), Some((0, 0)));
        assert_eq!(parse_cell_ref("B2"), Some((1, 1)));
        assert_eq!(parse_cell_ref("Z1"), Some((25, 0)));
        assert_eq!(parse_cell_ref("AA1"), Some((26, 0)));
    }

    #[test]
    fn test_parse_cell_range() {
        assert_eq!(parse_cell_range("A1:B2"), Some((0, 0, 1, 1)));
        assert_eq!(parse_cell_range("A1:C3"), Some((0, 0, 2, 2)));
    }
}
