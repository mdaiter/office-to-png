//! DOCX document parsing and rendering.
//!
//! Extracts styled content from Word documents for Canvas 2D rendering.

use crate::renderer::{points_to_pixels, Color, RenderBackend};
use crate::text_layout::{Paragraph, Rect, TextAlign, TextRun, TextStyle};
use docx_rs::*;
use image::GenericImageView;
use serde::{Deserialize, Serialize};

/// Error type for DOCX operations.
#[derive(Debug)]
pub enum DocxError {
    ParseError(String),
    RenderError(String),
}

impl std::fmt::Display for DocxError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            DocxError::ParseError(msg) => write!(f, "Parse error: {}", msg),
            DocxError::RenderError(msg) => write!(f, "Render error: {}", msg),
        }
    }
}

impl std::error::Error for DocxError {}

/// A styled text run extracted from DOCX.
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct StyledRun {
    pub text: String,
    pub font_family: String,
    pub font_size: f32,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strikethrough: bool,
    pub color: String,
    pub highlight: Option<String>,
}

impl Default for StyledRun {
    fn default() -> Self {
        Self {
            text: String::new(),
            font_family: "Calibri".to_string(),
            font_size: 11.0,
            bold: false,
            italic: false,
            underline: false,
            strikethrough: false,
            color: "#000000".to_string(),
            highlight: None,
        }
    }
}

impl StyledRun {
    pub fn to_text_style(&self) -> TextStyle {
        let mut style = TextStyle::new(&self.font_family, self.font_size);
        style.bold = self.bold;
        style.italic = self.italic;
        style.underline = self.underline;
        style.strikethrough = self.strikethrough;
        if let Some(color) = Color::from_hex(&self.color) {
            style.color = color.to_rgba_array();
        }
        if let Some(ref highlight) = self.highlight {
            if let Some(color) = Color::from_hex(highlight) {
                style.background = Some(color.to_rgba_array());
            }
        }
        style
    }

    pub fn to_text_run(&self) -> TextRun {
        TextRun::new(&self.text, self.to_text_style())
    }
}

/// A styled paragraph extracted from DOCX.
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct StyledParagraph {
    pub runs: Vec<StyledRun>,
    pub alignment: String,
    pub line_spacing: f32,
    pub space_before: f32,
    pub space_after: f32,
    pub first_line_indent: f32,
    pub left_indent: f32,
    pub right_indent: f32,
    pub is_heading: bool,
    pub heading_level: Option<u8>,
}

impl Default for StyledParagraph {
    fn default() -> Self {
        Self {
            runs: Vec::new(),
            alignment: "left".to_string(),
            line_spacing: 1.15,
            space_before: 0.0,
            space_after: 8.0,
            first_line_indent: 0.0,
            left_indent: 0.0,
            right_indent: 0.0,
            is_heading: false,
            heading_level: None,
        }
    }
}

impl StyledParagraph {
    pub fn to_paragraph(&self) -> Paragraph {
        let runs: Vec<TextRun> = self.runs.iter().map(|r| r.to_text_run()).collect();
        let align = match self.alignment.as_str() {
            "center" => TextAlign::Center,
            "right" => TextAlign::Right,
            "justify" | "both" => TextAlign::Justify,
            _ => TextAlign::Left,
        };

        Paragraph {
            runs,
            align,
            line_spacing: self.line_spacing,
            space_before: self.space_before,
            space_after: self.space_after,
            first_line_indent: self.first_line_indent,
            left_indent: self.left_indent,
            right_indent: self.right_indent,
        }
    }

    pub fn get_text(&self) -> String {
        self.runs.iter().map(|r| r.text.as_str()).collect()
    }
}

/// Table cell from DOCX
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct TableCell {
    pub paragraphs: Vec<StyledParagraph>,
    pub col_span: u32,
    pub width: Option<f32>,
    pub background: Option<String>,
}

/// Table row from DOCX
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct TableRow {
    pub cells: Vec<TableCell>,
}

/// Table from DOCX
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct Table {
    pub rows: Vec<TableRow>,
    pub column_widths: Vec<f32>,
}

/// An image from DOCX
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct DocxImage {
    /// Image ID (relationship ID)
    pub id: String,
    /// Width in points
    pub width: f32,
    /// Height in points
    pub height: f32,
    /// X position offset in points (for floating images)
    pub x: f32,
    /// Y position offset in points (for floating images)
    pub y: f32,
    /// Whether the image is inline (flows with text) or floating
    pub is_inline: bool,
    /// Image data (PNG, JPEG, etc.)
    #[serde(skip)]
    pub data: Vec<u8>,
    /// Image format (png, jpeg, etc.)
    pub format: String,
}

impl Default for DocxImage {
    fn default() -> Self {
        Self {
            id: String::new(),
            width: 100.0,
            height: 100.0,
            x: 0.0,
            y: 0.0,
            is_inline: true,
            data: Vec::new(),
            format: "png".to_string(),
        }
    }
}

/// Document element (paragraph, table, image, etc.)
#[derive(Clone, Debug, Serialize, Deserialize)]
pub enum DocumentElement {
    Paragraph(StyledParagraph),
    Table(Table),
    Image(DocxImage),
    PageBreak,
}

/// A parsed DOCX document with full styling.
pub struct DocxDocument {
    /// Document elements in order
    elements: Vec<DocumentElement>,
    /// Document title from properties
    title: Option<String>,
    /// Page width in points
    page_width: f32,
    /// Page height in points
    page_height: f32,
    /// Page margins
    margin_top: f32,
    margin_right: f32,
    margin_bottom: f32,
    margin_left: f32,
    /// Media files (images) keyed by relationship ID
    media: std::collections::HashMap<String, (Vec<u8>, String)>,
}

impl DocxDocument {
    /// Parse a DOCX document from bytes.
    pub fn from_bytes(data: &[u8]) -> Result<Self, DocxError> {
        let docx = read_docx(data)
            .map_err(|e| DocxError::ParseError(format!("Failed to parse DOCX: {:?}", e)))?;

        // Default page size (US Letter)
        let page_width = 612.0;
        let page_height = 792.0;
        let margin_top = 72.0;
        let margin_right = 72.0;
        let margin_bottom = 72.0;
        let margin_left = 72.0;

        // Use default page size (US Letter) - section properties parsing is complex
        // The docx-rs API uses builder patterns that make reading values tricky
        // For now, use sensible defaults
        let _ = &docx.document.section_property; // Acknowledge we have it

        // Build media map from document relationships and media files
        let mut media = std::collections::HashMap::new();

        // Extract media files from docx
        for (filename, data) in &docx.media {
            // filename is like "image1.png" - extract the relationship ID
            let format = filename.rsplit('.').next().unwrap_or("png").to_lowercase();
            // We'll map by filename for now since we need to match with relationship targets
            media.insert(filename.clone(), (data.clone(), format));
        }

        // Also check the images field which has (id, path, image, png) tuples
        for (id, path, _image, png) in &docx.images {
            // Extract format from path
            let format = path.rsplit('.').next().unwrap_or("png").to_lowercase();
            // Map by relationship ID
            media.insert(id.clone(), (png.0.clone(), format));
        }

        // Extract document elements
        let mut elements = Vec::new();

        for child in docx.document.children.iter() {
            match child {
                DocumentChild::Paragraph(para) => {
                    // Check for images in the paragraph
                    let images = Self::extract_images_from_paragraph(para, &media);
                    for img in images {
                        elements.push(DocumentElement::Image(img));
                    }

                    let styled = Self::extract_paragraph(para);
                    // Only add paragraph if it has content
                    if !styled.runs.is_empty() || styled.is_heading {
                        elements.push(DocumentElement::Paragraph(styled));
                    }
                }
                DocumentChild::Table(table) => {
                    let styled = Self::extract_table(table);
                    elements.push(DocumentElement::Table(styled));
                }
                _ => {}
            }
        }

        Ok(Self {
            elements,
            title: None,
            page_width,
            page_height,
            margin_top,
            margin_right,
            margin_bottom,
            margin_left,
            media,
        })
    }

    /// Extract images from a paragraph
    fn extract_images_from_paragraph(
        para: &docx_rs::Paragraph,
        media: &std::collections::HashMap<String, (Vec<u8>, String)>,
    ) -> Vec<DocxImage> {
        let mut images = Vec::new();

        for child in &para.children {
            if let ParagraphChild::Run(run) = child {
                for run_child in &run.children {
                    if let RunChild::Drawing(drawing) = run_child {
                        if let Some(img) = Self::extract_image_from_drawing(drawing, media) {
                            images.push(img);
                        }
                    }
                }
            }
        }

        images
    }

    /// Extract image data from a Drawing element
    fn extract_image_from_drawing(
        drawing: &Drawing,
        media: &std::collections::HashMap<String, (Vec<u8>, String)>,
    ) -> Option<DocxImage> {
        if let Some(DrawingData::Pic(pic)) = &drawing.data {
            // Convert EMU (English Metric Units) to points
            // 1 inch = 914400 EMU, 1 inch = 72 points
            // So 1 EMU = 72/914400 points
            const EMU_TO_POINTS: f32 = 72.0 / 914400.0;

            let width = pic.size.0 as f32 * EMU_TO_POINTS;
            let height = pic.size.1 as f32 * EMU_TO_POINTS;

            // Check if inline or floating
            let is_inline = matches!(pic.position_type, DrawingPositionType::Inline { .. });

            // Get position for floating images
            let (x, y) = if !is_inline {
                let x = match pic.position_h {
                    DrawingPosition::Offset(emu) => emu as f32 * EMU_TO_POINTS,
                    _ => 0.0,
                };
                let y = match pic.position_v {
                    DrawingPosition::Offset(emu) => emu as f32 * EMU_TO_POINTS,
                    _ => 0.0,
                };
                (x, y)
            } else {
                (0.0, 0.0)
            };

            // Try to get image data from the pic.image field first
            let (data, format) = if !pic.image.is_empty() {
                let format = Self::detect_image_format(&pic.image);
                (pic.image.clone(), format)
            } else {
                // Try to find by relationship ID
                // The id in Pic is used to look up in relationships
                if let Some((data, format)) = media.get(&pic.id) {
                    (data.clone(), format.clone())
                } else {
                    // Try matching by common patterns
                    let mut found = None;
                    for (key, (data, format)) in media.iter() {
                        if key.contains("image") || key.starts_with("rId") {
                            found = Some((data.clone(), format.clone()));
                            break;
                        }
                    }
                    found.unwrap_or_else(|| (Vec::new(), "png".to_string()))
                }
            };

            if !data.is_empty() || width > 0.0 {
                return Some(DocxImage {
                    id: pic.id.clone(),
                    width,
                    height,
                    x,
                    y,
                    is_inline,
                    data,
                    format,
                });
            }
        }

        None
    }

    /// Detect image format from bytes
    fn detect_image_format(data: &[u8]) -> String {
        if data.len() < 8 {
            return "png".to_string();
        }

        // Check magic bytes
        if data.starts_with(&[0x89, 0x50, 0x4E, 0x47]) {
            "png".to_string()
        } else if data.starts_with(&[0xFF, 0xD8, 0xFF]) {
            "jpeg".to_string()
        } else if data.starts_with(b"GIF87a") || data.starts_with(b"GIF89a") {
            "gif".to_string()
        } else if data.starts_with(b"RIFF") && data.len() > 12 && &data[8..12] == b"WEBP" {
            "webp".to_string()
        } else if data.starts_with(&[0x42, 0x4D]) {
            "bmp".to_string()
        } else {
            "png".to_string()
        }
    }

    fn extract_paragraph(para: &docx_rs::Paragraph) -> StyledParagraph {
        let mut styled = StyledParagraph::default();

        // Extract paragraph properties
        let props = &para.property;

        // Alignment
        if let Some(ref align) = props.alignment {
            styled.alignment = format!("{:?}", align.val).to_lowercase();
        }

        // Indentation
        if let Some(ref indent) = props.indent {
            if let Some(left) = indent.start {
                styled.left_indent = left as f32 / 20.0;
            }
            if let Some(right) = indent.end {
                styled.right_indent = right as f32 / 20.0;
            }
            if let Some(first) = indent.first_line_chars {
                styled.first_line_indent = first as f32 / 100.0 * 11.0;
            }
        }

        // Spacing - docx-rs uses builder pattern, values not easily accessible
        // Use default line spacing for now
        let _ = &props.line_spacing;

        // Check for heading style
        if let Some(ref style) = props.style {
            let style_id = &style.val;
            if style_id.starts_with("Heading") || style_id.starts_with("heading") {
                styled.is_heading = true;
                styled.heading_level = style_id
                    .chars()
                    .last()
                    .and_then(|c| c.to_digit(10).map(|d| d as u8));
            }
        }

        // Extract runs with styling
        for child in &para.children {
            if let ParagraphChild::Run(run) = child {
                let styled_run = Self::extract_run(run);
                if !styled_run.text.is_empty() {
                    styled.runs.push(styled_run);
                }
            }
        }

        styled
    }

    fn extract_run(run: &docx_rs::Run) -> StyledRun {
        let mut styled = StyledRun::default();

        // Extract text
        for child in &run.children {
            match child {
                RunChild::Text(t) => {
                    styled.text.push_str(&t.text);
                }
                RunChild::Tab(_) => {
                    styled.text.push('\t');
                }
                RunChild::Break(_) => {
                    styled.text.push('\n');
                }
                _ => {}
            }
        }

        // Extract run properties - docx-rs uses private fields with builder pattern
        // We check for presence of properties using Option pattern
        let props = &run.run_property;

        // Bold - check if bold property exists
        styled.bold = props.bold.is_some();

        // Italic
        styled.italic = props.italic.is_some();

        // Underline
        styled.underline = props.underline.is_some();

        // Strike
        styled.strikethrough = props.strike.is_some();

        // Note: Font size, family, color extraction would need reflection or
        // a different approach with docx-rs. Using defaults for now.
        // The library is designed for document creation, not parsing.

        styled
    }

    fn extract_table(table: &docx_rs::Table) -> Table {
        let mut rows = Vec::new();
        let mut column_widths = Vec::new();

        // Get column widths from grid
        for width in &table.grid {
            column_widths.push(*width as f32 / 20.0);
        }

        for table_child in &table.rows {
            // table.rows contains TableChild which wraps TableRow
            let TableChild::TableRow(row) = table_child;
            let mut table_row = TableRow { cells: Vec::new() };

            // Each TableRow has cells wrapped in TableRowChild
            for row_child in &row.cells {
                let TableRowChild::TableCell(cell) = row_child;
                let mut table_cell = TableCell {
                    paragraphs: Vec::new(),
                    col_span: 1,
                    width: None,
                    background: None,
                };

                // Extract cell content
                for content in &cell.children {
                    if let TableCellContent::Paragraph(para) = content {
                        table_cell.paragraphs.push(Self::extract_paragraph(para));
                    }
                }

                table_row.cells.push(table_cell);
            }
            rows.push(table_row);
        }

        Table {
            rows,
            column_widths,
        }
    }

    /// Get the number of pages (estimated).
    pub fn page_count(&self) -> usize {
        let content_height = self.page_height - self.margin_top - self.margin_bottom;
        let mut total_height = 0.0;

        for element in &self.elements {
            match element {
                DocumentElement::Paragraph(para) => {
                    let lines = para.runs.iter().map(|r| r.text.len()).sum::<usize>() / 80 + 1;
                    total_height += (lines as f32 * 20.0 * para.line_spacing)
                        + para.space_before
                        + para.space_after;
                }
                DocumentElement::Table(table) => {
                    total_height += table.rows.len() as f32 * 30.0;
                }
                DocumentElement::Image(img) => {
                    if img.is_inline {
                        total_height += img.height + 10.0; // Add some spacing
                    }
                    // Floating images don't add to flow height
                }
                DocumentElement::PageBreak => {
                    total_height += content_height;
                }
            }
        }

        ((total_height / content_height).ceil() as usize).max(1)
    }

    /// Get the document title.
    pub fn title(&self) -> Option<&str> {
        self.title.as_deref()
    }

    /// Get all elements.
    pub fn elements(&self) -> &[DocumentElement] {
        &self.elements
    }

    /// Get page dimensions in pixels (at 96 DPI).
    pub fn page_dimensions(&self) -> (f32, f32) {
        (
            points_to_pixels(self.page_width),
            points_to_pixels(self.page_height),
        )
    }

    /// Get margins in pixels.
    pub fn margins(&self) -> (f32, f32, f32, f32) {
        (
            points_to_pixels(self.margin_top),
            points_to_pixels(self.margin_right),
            points_to_pixels(self.margin_bottom),
            points_to_pixels(self.margin_left),
        )
    }

    /// Get content area in pixels.
    pub fn content_area(&self) -> Rect {
        let (top, right, bottom, left) = self.margins();
        let (width, height) = self.page_dimensions();
        Rect::new(left, top, width - left - right, height - top - bottom)
    }

    /// Render document to a backend.
    pub fn render<R: RenderBackend>(&self, renderer: &R, page: usize) -> Result<(), DocxError> {
        renderer
            .clear(Color::WHITE)
            .map_err(|e| DocxError::RenderError(e))?;

        let content_area = self.content_area();
        let page_height = content_area.height;
        let page_start_y = page as f32 * page_height;
        let page_end_y = page_start_y + page_height;

        let mut current_y = 0.0;

        for element in &self.elements {
            let element_height = self.estimate_element_height(element, content_area.width);

            if current_y + element_height >= page_start_y && current_y < page_end_y {
                let render_y = content_area.y + (current_y - page_start_y);

                if render_y >= content_area.y && render_y < content_area.y + content_area.height {
                    match element {
                        DocumentElement::Paragraph(para) => {
                            self.render_paragraph(
                                renderer,
                                para,
                                content_area.x,
                                render_y,
                                content_area.width,
                            )?;
                        }
                        DocumentElement::Table(table) => {
                            self.render_table(
                                renderer,
                                table,
                                content_area.x,
                                render_y,
                                content_area.width,
                            )?;
                        }
                        DocumentElement::Image(img) => {
                            self.render_image(renderer, img, content_area.x, render_y)?;
                        }
                        DocumentElement::PageBreak => {}
                    }
                }
            }

            current_y += element_height;
        }

        Ok(())
    }

    fn estimate_element_height(&self, element: &DocumentElement, width: f32) -> f32 {
        match element {
            DocumentElement::Paragraph(para) => para.to_paragraph().estimate_height(width),
            DocumentElement::Table(table) => table.rows.len() as f32 * 30.0,
            DocumentElement::Image(img) => {
                if img.is_inline {
                    points_to_pixels(img.height) + 10.0
                } else {
                    0.0 // Floating images don't affect flow
                }
            }
            DocumentElement::PageBreak => self.page_height - self.margin_top - self.margin_bottom,
        }
    }

    fn render_paragraph<R: RenderBackend>(
        &self,
        renderer: &R,
        para: &StyledParagraph,
        x: f32,
        y: f32,
        width: f32,
    ) -> Result<f32, DocxError> {
        let mut current_y = y + points_to_pixels(para.space_before);
        let mut current_x = x + points_to_pixels(para.left_indent + para.first_line_indent);
        let max_x = x + width - points_to_pixels(para.right_indent);
        let line_height = para
            .runs
            .first()
            .map(|r| points_to_pixels(r.font_size) * para.line_spacing * 1.2)
            .unwrap_or(20.0);

        for run in &para.runs {
            let style = run.to_text_style();
            let words: Vec<&str> = run.text.split_whitespace().collect();

            for (i, word) in words.iter().enumerate() {
                let word_with_space = if i < words.len() - 1 {
                    format!("{} ", word)
                } else {
                    word.to_string()
                };

                let metrics = renderer
                    .measure_text(&word_with_space, &style)
                    .map_err(|e| DocxError::RenderError(e))?;

                if current_x + metrics.width > max_x
                    && current_x > x + points_to_pixels(para.left_indent)
                {
                    current_y += line_height;
                    current_x = x + points_to_pixels(para.left_indent);
                }

                renderer
                    .draw_text(
                        &word_with_space,
                        current_x,
                        current_y + line_height * 0.8,
                        &style,
                    )
                    .map_err(|e| DocxError::RenderError(e))?;

                current_x += metrics.width;
            }
        }

        current_y += line_height + points_to_pixels(para.space_after);
        Ok(current_y - y)
    }

    fn render_table<R: RenderBackend>(
        &self,
        renderer: &R,
        table: &Table,
        x: f32,
        y: f32,
        max_width: f32,
    ) -> Result<f32, DocxError> {
        let mut current_y = y;
        let row_height = 25.0;

        let total_width: f32 = table
            .column_widths
            .iter()
            .map(|w| points_to_pixels(*w))
            .sum();
        let scale = if total_width > 0.0 {
            max_width.min(total_width) / total_width
        } else {
            1.0
        };
        let scaled_widths: Vec<f32> = table
            .column_widths
            .iter()
            .map(|w| points_to_pixels(*w) * scale)
            .collect();

        for row in &table.rows {
            let mut current_x = x;

            for (col_idx, cell) in row.cells.iter().enumerate() {
                let cell_width =
                    scaled_widths.get(col_idx).copied().unwrap_or(100.0) * cell.col_span as f32;

                if let Some(ref bg) = cell.background {
                    if let Some(color) = Color::from_hex(bg) {
                        renderer
                            .fill_rect(current_x, current_y, cell_width, row_height, color)
                            .map_err(|e| DocxError::RenderError(e))?;
                    }
                }

                let border_color = Color::BLACK;
                renderer
                    .draw_line(
                        current_x,
                        current_y,
                        current_x + cell_width,
                        current_y,
                        border_color,
                        1.0,
                    )
                    .map_err(|e| DocxError::RenderError(e))?;
                renderer
                    .draw_line(
                        current_x + cell_width,
                        current_y,
                        current_x + cell_width,
                        current_y + row_height,
                        border_color,
                        1.0,
                    )
                    .map_err(|e| DocxError::RenderError(e))?;
                renderer
                    .draw_line(
                        current_x,
                        current_y + row_height,
                        current_x + cell_width,
                        current_y + row_height,
                        border_color,
                        1.0,
                    )
                    .map_err(|e| DocxError::RenderError(e))?;
                renderer
                    .draw_line(
                        current_x,
                        current_y,
                        current_x,
                        current_y + row_height,
                        border_color,
                        1.0,
                    )
                    .map_err(|e| DocxError::RenderError(e))?;

                let text: String = cell
                    .paragraphs
                    .iter()
                    .flat_map(|p| p.runs.iter())
                    .map(|r| r.text.as_str())
                    .collect();

                if !text.is_empty() {
                    let style = cell
                        .paragraphs
                        .first()
                        .and_then(|p| p.runs.first())
                        .map(|r| r.to_text_style())
                        .unwrap_or_default();

                    renderer.save().map_err(|e| DocxError::RenderError(e))?;
                    renderer
                        .clip(current_x, current_y, cell_width, row_height)
                        .map_err(|e| DocxError::RenderError(e))?;
                    renderer
                        .draw_text(&text, current_x + 4.0, current_y + row_height * 0.7, &style)
                        .map_err(|e| DocxError::RenderError(e))?;
                    renderer.restore().map_err(|e| DocxError::RenderError(e))?;
                }

                current_x += cell_width;
            }

            current_y += row_height;
        }

        Ok(current_y - y)
    }

    /// Render an image element.
    fn render_image<R: RenderBackend>(
        &self,
        renderer: &R,
        img: &DocxImage,
        x: f32,
        y: f32,
    ) -> Result<f32, DocxError> {
        // Skip if no data
        if img.data.is_empty() {
            return Ok(0.0);
        }

        // Decode image using the image crate
        let decoded = image::load_from_memory(&img.data)
            .map_err(|e| DocxError::RenderError(format!("Failed to decode image: {}", e)))?;

        // Convert to RGBA8
        let rgba = decoded.to_rgba8();
        let (img_width, img_height) = rgba.dimensions();
        let rgba_data = rgba.into_raw();

        // Calculate destination dimensions in pixels
        let dest_width = points_to_pixels(img.width);
        let dest_height = points_to_pixels(img.height);

        // Draw the image
        renderer
            .draw_image(
                &rgba_data,
                img_width,
                img_height,
                x,
                y,
                dest_width,
                dest_height,
            )
            .map_err(|e| DocxError::RenderError(e))?;

        // Return height consumed (for inline images)
        if img.is_inline {
            Ok(dest_height + 10.0) // Add some spacing
        } else {
            Ok(0.0) // Floating images don't affect text flow
        }
    }

    /// Get paragraph count.
    pub fn paragraph_count(&self) -> usize {
        self.elements
            .iter()
            .filter(|e| matches!(e, DocumentElement::Paragraph(_)))
            .count()
    }

    /// Get paragraph text by index.
    pub fn get_paragraph(&self, index: usize) -> Option<String> {
        self.elements
            .iter()
            .filter_map(|e| match e {
                DocumentElement::Paragraph(p) => Some(p.get_text()),
                _ => None,
            })
            .nth(index)
    }

    // ========== Worker API: Pre-computed Layout ==========

    /// Compute pre-laid-out render primitives for a page.
    /// This is the heavy operation designed to run in a Web Worker.
    /// Returns the page data and a vector of image bytes (transferred separately).
    pub fn get_page_data(
        &self,
        page: usize,
    ) -> Option<(crate::render_data::DocxPageData, Vec<Vec<u8>>)> {
        use crate::render_data::{
            DocxPageData, ImageFormat, ImageMetadata, RenderPrimitive, RenderedImage, RenderedLine,
            RenderedRect, RenderedText,
        };

        let content_area = self.content_area();
        let page_height = content_area.height;
        let page_start_y = page as f32 * page_height;
        let page_end_y = page_start_y + page_height;

        let mut primitives = Vec::new();
        let mut images_meta = Vec::new();
        let mut image_bytes = Vec::new();

        // Start with clear
        primitives.push(RenderPrimitive::Clear([1.0, 1.0, 1.0, 1.0]));

        let mut current_y = 0.0;

        for element in &self.elements {
            let element_height = self.estimate_element_height(element, content_area.width);

            // Check if element is visible on this page
            if current_y + element_height >= page_start_y && current_y < page_end_y {
                let render_y = content_area.y + (current_y - page_start_y);

                if render_y >= content_area.y - element_height
                    && render_y < content_area.y + content_area.height
                {
                    match element {
                        DocumentElement::Paragraph(para) => {
                            self.layout_paragraph(
                                para,
                                content_area.x,
                                render_y,
                                content_area.width,
                                &mut primitives,
                            );
                        }
                        DocumentElement::Table(table) => {
                            self.layout_table(
                                table,
                                content_area.x,
                                render_y,
                                content_area.width,
                                &mut primitives,
                            );
                        }
                        DocumentElement::Image(img) => {
                            if !img.data.is_empty() {
                                let image_index = images_meta.len();

                                // Decode image to get actual dimensions
                                let (src_width, src_height) =
                                    if let Ok(decoded) = image::load_from_memory(&img.data) {
                                        decoded.dimensions()
                                    } else {
                                        (img.width as u32, img.height as u32)
                                    };

                                images_meta.push(ImageMetadata {
                                    src_width,
                                    src_height,
                                    format: ImageFormat::from_str(&img.format),
                                });
                                image_bytes.push(img.data.clone());

                                primitives.push(RenderPrimitive::Image(RenderedImage {
                                    x: if img.is_inline { content_area.x } else { img.x },
                                    y: if img.is_inline {
                                        render_y
                                    } else {
                                        render_y + img.y
                                    },
                                    dest_width: points_to_pixels(img.width),
                                    dest_height: points_to_pixels(img.height),
                                    image_index,
                                }));
                            }
                        }
                        DocumentElement::PageBreak => {}
                    }
                }
            }

            current_y += element_height;
        }

        Some((
            DocxPageData {
                page_index: page,
                page_width: points_to_pixels(self.page_width),
                page_height: points_to_pixels(self.page_height),
                primitives,
                images: images_meta,
            },
            image_bytes,
        ))
    }

    /// Get image bytes for a specific page (for separate Transferable).
    pub fn get_image_bytes_for_page(&self, page: usize) -> Vec<Vec<u8>> {
        if let Some((_, bytes)) = self.get_page_data(page) {
            bytes
        } else {
            Vec::new()
        }
    }

    /// Layout a paragraph into render primitives without drawing.
    fn layout_paragraph(
        &self,
        para: &StyledParagraph,
        x: f32,
        y: f32,
        width: f32,
        primitives: &mut Vec<crate::render_data::RenderPrimitive>,
    ) {
        use crate::render_data::{RenderPrimitive, RenderedText};

        let mut current_y = y + points_to_pixels(para.space_before);
        let mut current_x = x + points_to_pixels(para.left_indent + para.first_line_indent);
        let max_x = x + width - points_to_pixels(para.right_indent);
        let line_height = para
            .runs
            .first()
            .map(|r| points_to_pixels(r.font_size) * para.line_spacing * 1.2)
            .unwrap_or(20.0);

        for run in &para.runs {
            let style = run.to_text_style();
            let words: Vec<&str> = run.text.split_whitespace().collect();

            for (i, word) in words.iter().enumerate() {
                let word_with_space = if i < words.len() - 1 {
                    format!("{} ", word)
                } else {
                    word.to_string()
                };

                // Estimate text width (simplified character-based)
                let char_width = points_to_pixels(run.font_size) * 0.5;
                let word_width = word_with_space.len() as f32 * char_width;

                // Wrap to next line if needed
                if current_x + word_width > max_x
                    && current_x > x + points_to_pixels(para.left_indent)
                {
                    current_y += line_height;
                    current_x = x + points_to_pixels(para.left_indent);
                }

                primitives.push(RenderPrimitive::Text(RenderedText {
                    x: current_x,
                    y: current_y + line_height * 0.8,
                    text: word_with_space,
                    font_family: run.font_family.clone(),
                    font_size: run.font_size,
                    bold: run.bold,
                    italic: run.italic,
                    underline: run.underline,
                    strikethrough: run.strikethrough,
                    color: style.color,
                    background: style.background,
                }));

                current_x += word_width;
            }
        }
    }

    /// Layout a table into render primitives without drawing.
    fn layout_table(
        &self,
        table: &Table,
        x: f32,
        y: f32,
        max_width: f32,
        primitives: &mut Vec<crate::render_data::RenderPrimitive>,
    ) {
        use crate::render_data::{RenderPrimitive, RenderedLine, RenderedRect, RenderedText};
        use crate::renderer::Color;

        let mut current_y = y;
        let row_height = 25.0;

        // Calculate scaled column widths
        let total_width: f32 = table
            .column_widths
            .iter()
            .map(|w| points_to_pixels(*w))
            .sum();
        let scale = if total_width > 0.0 {
            max_width.min(total_width) / total_width
        } else {
            1.0
        };
        let scaled_widths: Vec<f32> = table
            .column_widths
            .iter()
            .map(|w| points_to_pixels(*w) * scale)
            .collect();

        let border_color = [0.0, 0.0, 0.0, 1.0]; // Black

        for row in &table.rows {
            let mut current_x = x;

            for (col_idx, cell) in row.cells.iter().enumerate() {
                let cell_width =
                    scaled_widths.get(col_idx).copied().unwrap_or(100.0) * cell.col_span as f32;

                // Cell background
                if let Some(ref bg) = cell.background {
                    if let Some(color) = Color::from_hex(bg) {
                        primitives.push(RenderPrimitive::Rect(RenderedRect {
                            x: current_x,
                            y: current_y,
                            width: cell_width,
                            height: row_height,
                            fill: Some(color.to_rgba_array()),
                            stroke: None,
                        }));
                    }
                }

                // Cell borders
                // Top
                primitives.push(RenderPrimitive::Line(RenderedLine {
                    x1: current_x,
                    y1: current_y,
                    x2: current_x + cell_width,
                    y2: current_y,
                    color: border_color,
                    width: 1.0,
                }));
                // Right
                primitives.push(RenderPrimitive::Line(RenderedLine {
                    x1: current_x + cell_width,
                    y1: current_y,
                    x2: current_x + cell_width,
                    y2: current_y + row_height,
                    color: border_color,
                    width: 1.0,
                }));
                // Bottom
                primitives.push(RenderPrimitive::Line(RenderedLine {
                    x1: current_x,
                    y1: current_y + row_height,
                    x2: current_x + cell_width,
                    y2: current_y + row_height,
                    color: border_color,
                    width: 1.0,
                }));
                // Left
                primitives.push(RenderPrimitive::Line(RenderedLine {
                    x1: current_x,
                    y1: current_y,
                    x2: current_x,
                    y2: current_y + row_height,
                    color: border_color,
                    width: 1.0,
                }));

                // Cell text
                let text: String = cell
                    .paragraphs
                    .iter()
                    .flat_map(|p| p.runs.iter())
                    .map(|r| r.text.as_str())
                    .collect();

                if !text.is_empty() {
                    let style = cell
                        .paragraphs
                        .first()
                        .and_then(|p| p.runs.first())
                        .map(|r| r.to_text_style())
                        .unwrap_or_default();

                    // Clip text to cell
                    primitives.push(RenderPrimitive::Save);
                    primitives.push(RenderPrimitive::Clip {
                        x: current_x,
                        y: current_y,
                        width: cell_width,
                        height: row_height,
                    });
                    primitives.push(RenderPrimitive::Text(RenderedText {
                        x: current_x + 4.0,
                        y: current_y + row_height * 0.7,
                        text,
                        font_family: style.font_family,
                        font_size: style.font_size,
                        bold: style.bold,
                        italic: style.italic,
                        underline: style.underline,
                        strikethrough: style.strikethrough,
                        color: style.color,
                        background: style.background,
                    }));
                    primitives.push(RenderPrimitive::Restore);
                }

                current_x += cell_width;
            }

            current_y += row_height;
        }
    }
}

#[allow(dead_code)]
fn highlight_to_hex(highlight: &str) -> String {
    match highlight.to_lowercase().as_str() {
        "yellow" => "#FFFF00".to_string(),
        "green" => "#00FF00".to_string(),
        "cyan" => "#00FFFF".to_string(),
        "magenta" => "#FF00FF".to_string(),
        "blue" => "#0000FF".to_string(),
        "red" => "#FF0000".to_string(),
        "darkblue" => "#00008B".to_string(),
        "darkcyan" => "#008B8B".to_string(),
        "darkgreen" => "#006400".to_string(),
        "darkmagenta" => "#8B008B".to_string(),
        "darkred" => "#8B0000".to_string(),
        "darkyellow" => "#808000".to_string(),
        "darkgray" => "#A9A9A9".to_string(),
        "lightgray" => "#D3D3D3".to_string(),
        "black" => "#000000".to_string(),
        "white" => "#FFFFFF".to_string(),
        _ => "#FFFF00".to_string(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_highlight_to_hex() {
        assert_eq!(highlight_to_hex("yellow"), "#FFFF00");
        assert_eq!(highlight_to_hex("Yellow"), "#FFFF00");
        assert_eq!(highlight_to_hex("darkblue"), "#00008B");
    }

    #[test]
    fn test_styled_run_default() {
        let run = StyledRun::default();
        assert_eq!(run.font_family, "Calibri");
        assert_eq!(run.font_size, 11.0);
        assert!(!run.bold);
    }
}
