//! Text shaping and layout using cosmic-text.
//!
//! This module provides proper text shaping (ligatures, kerning, etc.)
//! and line breaking using the cosmic-text library with embedded fonts.

use crate::fonts::FontManager;
use crate::renderer::Color;
use crate::text_layout::TextStyle;
use cosmic_text::{Attrs, Buffer, Family, FontSystem, Metrics, Shaping, Style, SwashCache, Weight};
use std::cell::RefCell;

/// A shaped glyph ready for rendering
#[derive(Clone, Debug)]
pub struct ShapedGlyph {
    /// Glyph ID in the font
    pub glyph_id: u16,
    /// X position relative to the line start
    pub x: f32,
    /// Y position (baseline)
    pub y: f32,
    /// Glyph advance width
    pub advance: f32,
    /// Font size used
    pub font_size: f32,
    /// Text color
    pub color: Color,
}

/// A shaped line of text
#[derive(Clone, Debug)]
pub struct ShapedLine {
    /// Glyphs in this line
    pub glyphs: Vec<ShapedGlyph>,
    /// Line width in pixels
    pub width: f32,
    /// Line height in pixels
    pub height: f32,
    /// Baseline offset from top
    pub baseline: f32,
}

/// A fully shaped text block
#[derive(Clone, Debug)]
pub struct ShapedText {
    /// Lines of shaped glyphs
    pub lines: Vec<ShapedLine>,
    /// Total width
    pub width: f32,
    /// Total height
    pub height: f32,
}

/// Text shaper using cosmic-text
pub struct TextShaper {
    font_system: RefCell<FontSystem>,
    swash_cache: RefCell<SwashCache>,
    font_manager: FontManager,
}

impl Default for TextShaper {
    fn default() -> Self {
        Self::new()
    }
}

impl TextShaper {
    /// Create a new text shaper with embedded fonts
    pub fn new() -> Self {
        let mut font_system = FontSystem::new();

        // Load embedded fonts
        let font_manager = FontManager::new();

        // Add fonts to the system
        for font_data in crate::fonts::get_embedded_fonts() {
            font_system.db_mut().load_font_data(font_data.data.to_vec());
        }

        Self {
            font_system: RefCell::new(font_system),
            swash_cache: RefCell::new(SwashCache::new()),
            font_manager,
        }
    }

    /// Shape text with the given style and max width
    pub fn shape_text(&self, text: &str, style: &TextStyle, max_width: Option<f32>) -> ShapedText {
        let mut font_system = self.font_system.borrow_mut();

        // Create metrics from style
        let metrics = Metrics::new(style.font_size, style.font_size * 1.2);

        // Create a buffer for the text
        let mut buffer = Buffer::new(&mut font_system, metrics);

        // Set buffer width for line wrapping
        if let Some(width) = max_width {
            buffer.set_size(&mut font_system, Some(width), None);
        }

        // Convert our style to cosmic-text attrs
        let weight = if style.bold {
            Weight::BOLD
        } else {
            Weight::NORMAL
        };

        let font_style = if style.italic {
            Style::Italic
        } else {
            Style::Normal
        };

        let attrs = Attrs::new()
            .family(Family::Name(&style.font_family))
            .weight(weight)
            .style(font_style);

        // Set the text
        buffer.set_text(&mut font_system, text, attrs, Shaping::Advanced);

        // Shape the text
        buffer.shape_until_scroll(&mut font_system, true);

        // Extract shaped glyphs
        let color = Color::from_rgba_array(style.color);
        let mut lines = Vec::new();
        let mut total_height = 0.0f32;
        let mut max_width = 0.0f32;

        for run in buffer.layout_runs() {
            let mut line_glyphs = Vec::new();
            let mut line_width = 0.0f32;

            for glyph in run.glyphs.iter() {
                line_glyphs.push(ShapedGlyph {
                    glyph_id: glyph.glyph_id,
                    x: glyph.x,
                    y: run.line_y,
                    advance: glyph.w,
                    font_size: style.font_size,
                    color,
                });
                line_width = line_width.max(glyph.x + glyph.w);
            }

            let line_height = run.line_height;
            lines.push(ShapedLine {
                glyphs: line_glyphs,
                width: line_width,
                height: line_height,
                baseline: line_height * 0.8, // Approximate baseline
            });

            total_height = total_height.max(run.line_y + line_height);
            max_width = max_width.max(line_width);
        }

        ShapedText {
            lines,
            width: max_width,
            height: total_height,
        }
    }

    /// Measure text without full shaping (faster for simple measurements)
    pub fn measure_text(&self, text: &str, style: &TextStyle) -> (f32, f32) {
        let shaped = self.shape_text(text, style, None);
        (shaped.width, shaped.height)
    }

    /// Shape text with multiple styles (rich text)
    ///
    /// Each span in the input is shaped with its own style, then combined.
    pub fn shape_rich_text(
        &self,
        spans: &[(String, TextStyle)],
        max_width: Option<f32>,
    ) -> ShapedText {
        if spans.is_empty() {
            return ShapedText {
                lines: Vec::new(),
                width: 0.0,
                height: 0.0,
            };
        }

        let mut font_system = self.font_system.borrow_mut();

        // Use the first span's font size for line metrics
        let base_size = spans.first().map(|(_, s)| s.font_size).unwrap_or(12.0);
        let metrics = Metrics::new(base_size, base_size * 1.2);

        let mut buffer = Buffer::new(&mut font_system, metrics);

        if let Some(width) = max_width {
            buffer.set_size(&mut font_system, Some(width), None);
        }

        // Build rich text as individual spans with attributes
        let rich_text: Vec<(&str, Attrs)> = spans
            .iter()
            .map(|(text, style)| {
                let weight = if style.bold {
                    Weight::BOLD
                } else {
                    Weight::NORMAL
                };

                let font_style = if style.italic {
                    Style::Italic
                } else {
                    Style::Normal
                };

                let attrs = Attrs::new()
                    .family(Family::Name(&style.font_family))
                    .weight(weight)
                    .style(font_style);

                (text.as_str(), attrs)
            })
            .collect();

        buffer.set_rich_text(&mut font_system, rich_text, Attrs::new(), Shaping::Advanced);
        buffer.shape_until_scroll(&mut font_system, true);

        // Extract shaped glyphs
        let default_color = Color::BLACK;
        let mut lines = Vec::new();
        let mut total_height = 0.0f32;
        let mut max_line_width = 0.0f32;

        for run in buffer.layout_runs() {
            let mut line_glyphs = Vec::new();
            let mut line_width = 0.0f32;

            for glyph in run.glyphs.iter() {
                // TODO: Map glyph back to span for correct color
                line_glyphs.push(ShapedGlyph {
                    glyph_id: glyph.glyph_id,
                    x: glyph.x,
                    y: run.line_y,
                    advance: glyph.w,
                    font_size: base_size,
                    color: default_color,
                });
                line_width = line_width.max(glyph.x + glyph.w);
            }

            let line_height = run.line_height;
            lines.push(ShapedLine {
                glyphs: line_glyphs,
                width: line_width,
                height: line_height,
                baseline: line_height * 0.8,
            });

            total_height = total_height.max(run.line_y + line_height);
            max_line_width = max_line_width.max(line_width);
        }

        ShapedText {
            lines,
            width: max_line_width,
            height: total_height,
        }
    }

    /// Check if a font family is available
    pub fn has_font(&self, family: &str) -> bool {
        self.font_manager.has_font(family)
    }

    /// Get available font families
    pub fn available_fonts(&self) -> Vec<&str> {
        self.font_manager.available_fonts()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_text_shaper_creation() {
        let shaper = TextShaper::new();
        assert!(shaper.has_font("Noto Sans"));
    }

    #[test]
    fn test_shape_simple_text() {
        let shaper = TextShaper::new();
        let style = TextStyle::new("Noto Sans", 12.0);
        let shaped = shaper.shape_text("Hello", &style, None);

        // Should have at least one line
        assert!(!shaped.lines.is_empty());
        // Width should be positive
        assert!(shaped.width > 0.0);
    }

    #[test]
    fn test_measure_text() {
        let shaper = TextShaper::new();
        let style = TextStyle::new("Noto Sans", 12.0);
        let (width, height) = shaper.measure_text("Hello World", &style);

        assert!(width > 0.0);
        assert!(height > 0.0);
    }
}
