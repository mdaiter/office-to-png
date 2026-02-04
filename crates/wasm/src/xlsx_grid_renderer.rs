//! XLSX grid/cell rendering using Canvas 2D.
//!
//! Renders spreadsheet cells with full styling support.

use crate::renderer::{Color, RenderBackend};
use crate::text_layout::TextStyle;
use crate::xlsx_renderer::{SheetData, StyledCell};

/// Default cell padding in pixels
const CELL_PADDING: f32 = 4.0;

/// Grid line color
const GRID_COLOR: Color = Color {
    r: 0.82,
    g: 0.82,
    b: 0.82,
    a: 1.0,
};

/// Header background color
const HEADER_BG: Color = Color {
    r: 0.95,
    g: 0.95,
    b: 0.95,
    a: 1.0,
};

/// Selection color
#[allow(dead_code)]
const SELECTION_COLOR: Color = Color {
    r: 0.2,
    g: 0.4,
    b: 0.8,
    a: 0.3,
};

/// Renderer for Excel spreadsheets.
pub struct XlsxGridRenderer {
    /// Show row/column headers
    pub show_headers: bool,
    /// Header width (for row numbers)
    pub header_width: f32,
    /// Header height (for column letters)
    pub header_height: f32,
    /// Frozen rows (always visible at top)
    pub frozen_rows: u32,
    /// Frozen columns (always visible at left)
    pub frozen_cols: u32,
    /// Current scroll offset X
    pub scroll_x: f32,
    /// Current scroll offset Y
    pub scroll_y: f32,
    /// Zoom level (1.0 = 100%)
    pub zoom: f32,
}

impl Default for XlsxGridRenderer {
    fn default() -> Self {
        Self {
            show_headers: true,
            header_width: 50.0,
            header_height: 25.0,
            frozen_rows: 0,
            frozen_cols: 0,
            scroll_x: 0.0,
            scroll_y: 0.0,
            zoom: 1.0,
        }
    }
}

impl XlsxGridRenderer {
    pub fn new() -> Self {
        Self::default()
    }

    /// Set frozen rows and columns.
    pub fn set_frozen(&mut self, rows: u32, cols: u32) {
        self.frozen_rows = rows;
        self.frozen_cols = cols;
    }

    /// Set scroll position.
    pub fn set_scroll(&mut self, x: f32, y: f32) {
        self.scroll_x = x.max(0.0);
        self.scroll_y = y.max(0.0);
    }

    /// Set zoom level.
    pub fn set_zoom(&mut self, zoom: f32) {
        self.zoom = zoom.clamp(0.1, 5.0);
    }

    /// Render a sheet to the given backend.
    pub fn render<R: RenderBackend>(&self, renderer: &R, sheet: &SheetData) -> Result<(), String> {
        let width = renderer.width();
        let height = renderer.height();

        // Clear background
        renderer.clear(Color::WHITE)?;

        // Calculate content area
        let content_x = if self.show_headers {
            self.header_width
        } else {
            0.0
        };
        let content_y = if self.show_headers {
            self.header_height
        } else {
            0.0
        };
        let content_width = width - content_x;
        let content_height = height - content_y;

        // Draw row/column headers if enabled
        if self.show_headers {
            self.render_headers(renderer, sheet, width, height)?;
        }

        // Set up clipping for content area
        renderer.save()?;
        renderer.clip(content_x, content_y, content_width, content_height)?;

        // Render grid and cells
        self.render_grid(
            renderer,
            sheet,
            content_x,
            content_y,
            content_width,
            content_height,
        )?;
        self.render_cells(renderer, sheet, content_x, content_y)?;

        renderer.restore()?;

        Ok(())
    }

    fn render_headers<R: RenderBackend>(
        &self,
        renderer: &R,
        sheet: &SheetData,
        width: f32,
        height: f32,
    ) -> Result<(), String> {
        let header_style = TextStyle::new("Arial", 11.0);

        // Draw header backgrounds
        renderer.fill_rect(0.0, 0.0, self.header_width, height, HEADER_BG)?;
        renderer.fill_rect(0.0, 0.0, width, self.header_height, HEADER_BG)?;

        // Draw corner cell
        renderer.fill_rect(0.0, 0.0, self.header_width, self.header_height, HEADER_BG)?;
        renderer.stroke_rect(
            0.0,
            0.0,
            self.header_width,
            self.header_height,
            &crate::renderer::BorderStyle::default(),
        )?;

        // Draw column headers (A, B, C, ...)
        let mut x = self.header_width - self.scroll_x;
        for (col_idx, &col_width) in sheet.column_widths.iter().enumerate() {
            let scaled_width = col_width * self.zoom;

            if x + scaled_width > 0.0 && x < width {
                let col_letter = column_to_letter(col_idx as u32);

                // Center the text
                let metrics = renderer.measure_text(&col_letter, &header_style)?;
                let text_x = x + (scaled_width - metrics.width) / 2.0;
                let text_y = self.header_height / 2.0 + metrics.height / 3.0;

                renderer.draw_text(&col_letter, text_x, text_y, &header_style)?;

                // Draw border
                renderer.draw_line(
                    x + scaled_width,
                    0.0,
                    x + scaled_width,
                    self.header_height,
                    Color::from_hex("#c0c0c0").unwrap_or(GRID_COLOR),
                    1.0,
                )?;
            }

            x += scaled_width;
            if x > width {
                break;
            }
        }

        // Draw row headers (1, 2, 3, ...)
        let mut y = self.header_height - self.scroll_y;
        for (row_idx, &row_height) in sheet.row_heights.iter().enumerate() {
            let scaled_height = row_height * self.zoom;

            if y + scaled_height > 0.0 && y < height {
                let row_num = (row_idx + 1).to_string();

                // Center the text
                let metrics = renderer.measure_text(&row_num, &header_style)?;
                let text_x = (self.header_width - metrics.width) / 2.0;
                let text_y = y + scaled_height / 2.0 + metrics.height / 3.0;

                renderer.draw_text(&row_num, text_x, text_y, &header_style)?;

                // Draw border
                renderer.draw_line(
                    0.0,
                    y + scaled_height,
                    self.header_width,
                    y + scaled_height,
                    Color::from_hex("#c0c0c0").unwrap_or(GRID_COLOR),
                    1.0,
                )?;
            }

            y += scaled_height;
            if y > height {
                break;
            }
        }

        Ok(())
    }

    fn render_grid<R: RenderBackend>(
        &self,
        renderer: &R,
        sheet: &SheetData,
        offset_x: f32,
        offset_y: f32,
        width: f32,
        height: f32,
    ) -> Result<(), String> {
        // Draw vertical grid lines
        let mut x = offset_x - self.scroll_x;
        for &col_width in &sheet.column_widths {
            let scaled_width = col_width * self.zoom;
            x += scaled_width;

            if x > offset_x && x < offset_x + width {
                renderer.draw_line(x, offset_y, x, offset_y + height, GRID_COLOR, 1.0)?;
            }

            if x > offset_x + width {
                break;
            }
        }

        // Draw horizontal grid lines
        let mut y = offset_y - self.scroll_y;
        for &row_height in &sheet.row_heights {
            let scaled_height = row_height * self.zoom;
            y += scaled_height;

            if y > offset_y && y < offset_y + height {
                renderer.draw_line(offset_x, y, offset_x + width, y, GRID_COLOR, 1.0)?;
            }

            if y > offset_y + height {
                break;
            }
        }

        Ok(())
    }

    fn render_cells<R: RenderBackend>(
        &self,
        renderer: &R,
        sheet: &SheetData,
        offset_x: f32,
        offset_y: f32,
    ) -> Result<(), String> {
        // Pre-calculate row Y positions
        let mut row_y_positions = Vec::with_capacity(sheet.row_heights.len() + 1);
        let mut y = offset_y - self.scroll_y;
        row_y_positions.push(y);
        for &row_height in &sheet.row_heights {
            y += row_height * self.zoom;
            row_y_positions.push(y);
        }

        // Pre-calculate column X positions
        let mut col_x_positions = Vec::with_capacity(sheet.column_widths.len() + 1);
        let mut x = offset_x - self.scroll_x;
        col_x_positions.push(x);
        for &col_width in &sheet.column_widths {
            x += col_width * self.zoom;
            col_x_positions.push(x);
        }

        // Render each cell
        for cell in &sheet.cells {
            let row = cell.row as usize;
            let col = cell.col as usize;

            if row >= row_y_positions.len() - 1 || col >= col_x_positions.len() - 1 {
                continue;
            }

            let cell_x = col_x_positions[col];
            let cell_y = row_y_positions[row];
            let cell_width = col_x_positions.get(col + 1).copied().unwrap_or(cell_x) - cell_x;
            let cell_height = row_y_positions.get(row + 1).copied().unwrap_or(cell_y) - cell_y;

            // Skip cells outside visible area
            if cell_x + cell_width < offset_x || cell_y + cell_height < offset_y {
                continue;
            }

            self.render_cell(renderer, cell, cell_x, cell_y, cell_width, cell_height)?;
        }

        Ok(())
    }

    fn render_cell<R: RenderBackend>(
        &self,
        renderer: &R,
        cell: &StyledCell,
        x: f32,
        y: f32,
        width: f32,
        height: f32,
    ) -> Result<(), String> {
        // Draw background
        if !cell.bg_color.is_empty() {
            if let Some(color) = Color::from_hex(&cell.bg_color) {
                renderer.fill_rect(x, y, width, height, color)?;
            }
        }

        // Draw borders
        if let Some(ref border) = cell.border_top {
            if let Some(color) = Color::from_hex(&border.color) {
                renderer.draw_line(x, y, x + width, y, color, border.width)?;
            }
        }
        if let Some(ref border) = cell.border_right {
            if let Some(color) = Color::from_hex(&border.color) {
                renderer.draw_line(x + width, y, x + width, y + height, color, border.width)?;
            }
        }
        if let Some(ref border) = cell.border_bottom {
            if let Some(color) = Color::from_hex(&border.color) {
                renderer.draw_line(x, y + height, x + width, y + height, color, border.width)?;
            }
        }
        if let Some(ref border) = cell.border_left {
            if let Some(color) = Color::from_hex(&border.color) {
                renderer.draw_line(x, y, x, y + height, color, border.width)?;
            }
        }

        // Draw text
        if !cell.value.is_empty() {
            let style = self.cell_to_text_style(cell);
            let metrics = renderer.measure_text(&cell.value, &style)?;

            // Calculate text position based on alignment
            let text_x = match cell.h_align.as_str() {
                "center" => x + (width - metrics.width) / 2.0,
                "right" => x + width - metrics.width - CELL_PADDING,
                _ => x + CELL_PADDING, // left or general
            };

            let text_y = match cell.v_align.as_str() {
                "top" => y + metrics.ascent + CELL_PADDING,
                "bottom" => y + height - metrics.descent - CELL_PADDING,
                _ => y + height / 2.0 + metrics.height / 3.0, // center
            };

            // Clip text to cell bounds
            renderer.save()?;
            renderer.clip(x, y, width, height)?;
            renderer.draw_text(&cell.value, text_x, text_y, &style)?;
            renderer.restore()?;
        }

        Ok(())
    }

    fn cell_to_text_style(&self, cell: &StyledCell) -> TextStyle {
        let mut style = TextStyle::new(&cell.font_family, cell.font_size * self.zoom);
        style.bold = cell.bold;
        style.italic = cell.italic;
        style.underline = cell.underline;

        if let Some(color) = Color::from_hex(&cell.text_color) {
            style.color = color.to_rgba_array();
        }

        style
    }

    /// Get the cell at a given pixel position.
    pub fn cell_at_position(&self, sheet: &SheetData, x: f32, y: f32) -> Option<(u32, u32)> {
        let offset_x = if self.show_headers {
            self.header_width
        } else {
            0.0
        };
        let offset_y = if self.show_headers {
            self.header_height
        } else {
            0.0
        };

        if x < offset_x || y < offset_y {
            return None; // In header area
        }

        // Find column
        let mut col_x = offset_x - self.scroll_x;
        let mut col = 0u32;
        for (idx, &width) in sheet.column_widths.iter().enumerate() {
            let next_x = col_x + width * self.zoom;
            if x >= col_x && x < next_x {
                col = idx as u32;
                break;
            }
            col_x = next_x;
        }

        // Find row
        let mut row_y = offset_y - self.scroll_y;
        let mut row = 0u32;
        for (idx, &height) in sheet.row_heights.iter().enumerate() {
            let next_y = row_y + height * self.zoom;
            if y >= row_y && y < next_y {
                row = idx as u32;
                break;
            }
            row_y = next_y;
        }

        Some((row, col))
    }

    /// Get the total content size.
    pub fn content_size(&self, sheet: &SheetData) -> (f32, f32) {
        let width: f32 = sheet.column_widths.iter().sum::<f32>() * self.zoom;
        let height: f32 = sheet.row_heights.iter().sum::<f32>() * self.zoom;
        (width, height)
    }
}

/// Convert column index to Excel-style letter (0 -> A, 25 -> Z, 26 -> AA, etc.)
fn column_to_letter(col: u32) -> String {
    let mut result = String::new();
    let mut n = col + 1;

    while n > 0 {
        n -= 1;
        result.insert(0, ((n % 26) as u8 + b'A') as char);
        n /= 26;
    }

    result
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_column_to_letter() {
        assert_eq!(column_to_letter(0), "A");
        assert_eq!(column_to_letter(1), "B");
        assert_eq!(column_to_letter(25), "Z");
        assert_eq!(column_to_letter(26), "AA");
        assert_eq!(column_to_letter(27), "AB");
        assert_eq!(column_to_letter(701), "ZZ");
        assert_eq!(column_to_letter(702), "AAA");
    }
}
