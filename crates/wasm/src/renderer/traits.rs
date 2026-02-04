//! Traits and types for rendering backends.

use crate::text_layout::TextStyle;

/// RGBA color representation.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct Color {
    pub r: f32,
    pub g: f32,
    pub b: f32,
    pub a: f32,
}

impl Color {
    pub const BLACK: Color = Color {
        r: 0.0,
        g: 0.0,
        b: 0.0,
        a: 1.0,
    };
    pub const WHITE: Color = Color {
        r: 1.0,
        g: 1.0,
        b: 1.0,
        a: 1.0,
    };
    pub const TRANSPARENT: Color = Color {
        r: 0.0,
        g: 0.0,
        b: 0.0,
        a: 0.0,
    };

    pub fn new(r: f32, g: f32, b: f32, a: f32) -> Self {
        Self { r, g, b, a }
    }

    pub fn rgb(r: f32, g: f32, b: f32) -> Self {
        Self { r, g, b, a: 1.0 }
    }

    pub fn from_rgba_array(arr: [f32; 4]) -> Self {
        Self {
            r: arr[0],
            g: arr[1],
            b: arr[2],
            a: arr[3],
        }
    }

    pub fn to_rgba_array(&self) -> [f32; 4] {
        [self.r, self.g, self.b, self.a]
    }

    /// Convert to CSS rgba() string
    pub fn to_css(&self) -> String {
        format!(
            "rgba({}, {}, {}, {})",
            (self.r * 255.0) as u8,
            (self.g * 255.0) as u8,
            (self.b * 255.0) as u8,
            self.a
        )
    }

    /// Parse from hex string (#RGB, #RRGGBB, or #RRGGBBAA)
    pub fn from_hex(hex: &str) -> Option<Self> {
        let hex = hex.trim_start_matches('#');

        let (r, g, b, a) = match hex.len() {
            3 => {
                let r = u8::from_str_radix(&hex[0..1], 16).ok()? * 17;
                let g = u8::from_str_radix(&hex[1..2], 16).ok()? * 17;
                let b = u8::from_str_radix(&hex[2..3], 16).ok()? * 17;
                (r, g, b, 255)
            }
            6 => {
                let r = u8::from_str_radix(&hex[0..2], 16).ok()?;
                let g = u8::from_str_radix(&hex[2..4], 16).ok()?;
                let b = u8::from_str_radix(&hex[4..6], 16).ok()?;
                (r, g, b, 255)
            }
            8 => {
                let r = u8::from_str_radix(&hex[0..2], 16).ok()?;
                let g = u8::from_str_radix(&hex[2..4], 16).ok()?;
                let b = u8::from_str_radix(&hex[4..6], 16).ok()?;
                let a = u8::from_str_radix(&hex[6..8], 16).ok()?;
                (r, g, b, a)
            }
            _ => return None,
        };

        Some(Self {
            r: r as f32 / 255.0,
            g: g as f32 / 255.0,
            b: b as f32 / 255.0,
            a: a as f32 / 255.0,
        })
    }
}

impl Default for Color {
    fn default() -> Self {
        Self::BLACK
    }
}

/// Text measurement results.
#[derive(Clone, Debug)]
pub struct TextMetrics {
    /// Width of the text in pixels
    pub width: f32,
    /// Height of the text (approximate, based on font size)
    pub height: f32,
    /// Ascent (distance from baseline to top)
    pub ascent: f32,
    /// Descent (distance from baseline to bottom)
    pub descent: f32,
}

/// Border style for drawing.
#[derive(Clone, Debug)]
pub struct BorderStyle {
    pub width: f32,
    pub color: Color,
    pub dash_pattern: Option<Vec<f32>>,
}

impl Default for BorderStyle {
    fn default() -> Self {
        Self {
            width: 1.0,
            color: Color::BLACK,
            dash_pattern: None,
        }
    }
}

/// A draw command for batched rendering.
#[derive(Clone, Debug)]
pub enum DrawCommand {
    /// Clear the canvas with a color
    Clear(Color),

    /// Draw a filled rectangle
    FillRect {
        x: f32,
        y: f32,
        width: f32,
        height: f32,
        color: Color,
    },

    /// Draw a stroked rectangle
    StrokeRect {
        x: f32,
        y: f32,
        width: f32,
        height: f32,
        border: BorderStyle,
    },

    /// Draw text
    DrawText {
        text: String,
        x: f32,
        y: f32,
        style: TextStyle,
    },

    /// Draw a line
    DrawLine {
        x1: f32,
        y1: f32,
        x2: f32,
        y2: f32,
        color: Color,
        width: f32,
    },

    /// Draw an image (raw RGBA data)
    DrawImage {
        data: Vec<u8>,
        width: u32,
        height: u32,
        x: f32,
        y: f32,
        dest_width: f32,
        dest_height: f32,
    },

    /// Save the current state (transforms, clip, etc.)
    Save,

    /// Restore a previously saved state
    Restore,

    /// Apply a translation transform
    Translate { x: f32, y: f32 },

    /// Apply a scale transform
    Scale { x: f32, y: f32 },

    /// Set clipping rectangle
    Clip {
        x: f32,
        y: f32,
        width: f32,
        height: f32,
    },
}

/// Trait for rendering backends.
pub trait RenderBackend {
    /// Get the canvas width
    fn width(&self) -> f32;

    /// Get the canvas height
    fn height(&self) -> f32;

    /// Resize the canvas
    fn resize(&mut self, width: f32, height: f32) -> Result<(), String>;

    /// Clear the canvas with a color
    fn clear(&self, color: Color) -> Result<(), String>;

    /// Draw a filled rectangle
    fn fill_rect(
        &self,
        x: f32,
        y: f32,
        width: f32,
        height: f32,
        color: Color,
    ) -> Result<(), String>;

    /// Draw a stroked rectangle
    fn stroke_rect(
        &self,
        x: f32,
        y: f32,
        width: f32,
        height: f32,
        border: &BorderStyle,
    ) -> Result<(), String>;

    /// Draw text at a position (baseline is at y)
    fn draw_text(&self, text: &str, x: f32, y: f32, style: &TextStyle) -> Result<(), String>;

    /// Measure text dimensions
    fn measure_text(&self, text: &str, style: &TextStyle) -> Result<TextMetrics, String>;

    /// Draw a line
    fn draw_line(
        &self,
        x1: f32,
        y1: f32,
        x2: f32,
        y2: f32,
        color: Color,
        width: f32,
    ) -> Result<(), String>;

    /// Draw an image from raw RGBA data
    fn draw_image(
        &self,
        data: &[u8],
        img_width: u32,
        img_height: u32,
        x: f32,
        y: f32,
        dest_width: f32,
        dest_height: f32,
    ) -> Result<(), String>;

    /// Save the current drawing state
    fn save(&self) -> Result<(), String>;

    /// Restore a previously saved state
    fn restore(&self) -> Result<(), String>;

    /// Apply a translation
    fn translate(&self, x: f32, y: f32) -> Result<(), String>;

    /// Apply a scale
    fn scale(&self, x: f32, y: f32) -> Result<(), String>;

    /// Set clipping rectangle
    fn clip(&self, x: f32, y: f32, width: f32, height: f32) -> Result<(), String>;

    /// Execute a batch of draw commands
    fn execute_commands(&self, commands: &[DrawCommand]) -> Result<(), String> {
        for cmd in commands {
            match cmd {
                DrawCommand::Clear(color) => self.clear(*color)?,
                DrawCommand::FillRect {
                    x,
                    y,
                    width,
                    height,
                    color,
                } => {
                    self.fill_rect(*x, *y, *width, *height, *color)?;
                }
                DrawCommand::StrokeRect {
                    x,
                    y,
                    width,
                    height,
                    border,
                } => {
                    self.stroke_rect(*x, *y, *width, *height, border)?;
                }
                DrawCommand::DrawText { text, x, y, style } => {
                    self.draw_text(text, *x, *y, style)?;
                }
                DrawCommand::DrawLine {
                    x1,
                    y1,
                    x2,
                    y2,
                    color,
                    width,
                } => {
                    self.draw_line(*x1, *y1, *x2, *y2, *color, *width)?;
                }
                DrawCommand::DrawImage {
                    data,
                    width,
                    height,
                    x,
                    y,
                    dest_width,
                    dest_height,
                } => {
                    self.draw_image(data, *width, *height, *x, *y, *dest_width, *dest_height)?;
                }
                DrawCommand::Save => self.save()?,
                DrawCommand::Restore => self.restore()?,
                DrawCommand::Translate { x, y } => self.translate(*x, *y)?,
                DrawCommand::Scale { x, y } => self.scale(*x, *y)?,
                DrawCommand::Clip {
                    x,
                    y,
                    width,
                    height,
                } => self.clip(*x, *y, *width, *height)?,
            }
        }
        Ok(())
    }

    /// Export the canvas to PNG bytes
    fn export_png(&self) -> Result<Vec<u8>, String>;
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_color_from_hex() {
        assert_eq!(Color::from_hex("#000000"), Some(Color::BLACK));
        assert_eq!(Color::from_hex("#ffffff"), Some(Color::WHITE));
        assert_eq!(Color::from_hex("fff"), Some(Color::WHITE));
        assert_eq!(Color::from_hex("#ff0000"), Some(Color::rgb(1.0, 0.0, 0.0)));
    }

    #[test]
    fn test_color_to_css() {
        assert_eq!(Color::BLACK.to_css(), "rgba(0, 0, 0, 1)");
        assert_eq!(Color::WHITE.to_css(), "rgba(255, 255, 255, 1)");
    }
}
