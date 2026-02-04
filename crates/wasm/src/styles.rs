//! Style definitions for document rendering.

use crate::text_layout::TextStyle;

/// Border style for cells and shapes.
#[derive(Clone, Debug)]
pub struct BorderStyle {
    /// Border width in points
    pub width: f32,
    /// Border color (RGBA)
    pub color: [f32; 4],
    /// Border style type
    pub style: BorderType,
}

impl Default for BorderStyle {
    fn default() -> Self {
        Self {
            width: 1.0,
            color: [0.0, 0.0, 0.0, 1.0], // Black
            style: BorderType::Solid,
        }
    }
}

/// Border type/pattern.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum BorderType {
    None,
    Solid,
    Dashed,
    Dotted,
    Double,
}

/// Fill style for shapes and cells.
#[derive(Clone, Debug)]
pub enum FillStyle {
    /// No fill (transparent)
    None,
    /// Solid color fill
    Solid([f32; 4]),
    /// Linear gradient
    LinearGradient {
        start: [f32; 4],
        end: [f32; 4],
        angle: f32,
    },
    /// Pattern fill
    Pattern {
        foreground: [f32; 4],
        background: [f32; 4],
        pattern: PatternType,
    },
}

impl Default for FillStyle {
    fn default() -> Self {
        FillStyle::None
    }
}

/// Pattern types for pattern fills.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum PatternType {
    Solid,
    Gray125,
    Gray0625,
    DarkDown,
    DarkGray,
    DarkGrid,
    DarkHorizontal,
    DarkTrellis,
    DarkUp,
    DarkVertical,
    LightDown,
    LightGray,
    LightGrid,
    LightHorizontal,
    LightTrellis,
    LightUp,
    LightVertical,
    MediumGray,
}

/// Cell style for spreadsheets.
#[derive(Clone, Debug, Default)]
pub struct CellStyle {
    /// Text style
    pub text: TextStyle,
    /// Fill style
    pub fill: FillStyle,
    /// Top border
    pub border_top: Option<BorderStyle>,
    /// Right border
    pub border_right: Option<BorderStyle>,
    /// Bottom border
    pub border_bottom: Option<BorderStyle>,
    /// Left border
    pub border_left: Option<BorderStyle>,
    /// Horizontal alignment
    pub h_align: HorizontalAlign,
    /// Vertical alignment
    pub v_align: VerticalAlign,
    /// Text wrapping
    pub wrap_text: bool,
    /// Number format (for display)
    pub number_format: Option<String>,
}

/// Horizontal alignment options.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Default)]
pub enum HorizontalAlign {
    #[default]
    General,
    Left,
    Center,
    Right,
    Fill,
    Justify,
    CenterContinuous,
    Distributed,
}

/// Vertical alignment options.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Default)]
pub enum VerticalAlign {
    Top,
    #[default]
    Center,
    Bottom,
    Justify,
    Distributed,
}

/// Convert a hex color string to RGBA.
pub fn hex_to_rgba(hex: &str) -> Option<[f32; 4]> {
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

    Some([
        r as f32 / 255.0,
        g as f32 / 255.0,
        b as f32 / 255.0,
        a as f32 / 255.0,
    ])
}

/// Convert RGBA to hex color string.
pub fn rgba_to_hex(color: [f32; 4]) -> String {
    let r = (color[0] * 255.0) as u8;
    let g = (color[1] * 255.0) as u8;
    let b = (color[2] * 255.0) as u8;
    let a = (color[3] * 255.0) as u8;

    if a == 255 {
        format!("#{:02x}{:02x}{:02x}", r, g, b)
    } else {
        format!("#{:02x}{:02x}{:02x}{:02x}", r, g, b, a)
    }
}

/// Standard colors for Office documents.
pub mod colors {
    pub const BLACK: [f32; 4] = [0.0, 0.0, 0.0, 1.0];
    pub const WHITE: [f32; 4] = [1.0, 1.0, 1.0, 1.0];
    pub const RED: [f32; 4] = [1.0, 0.0, 0.0, 1.0];
    pub const GREEN: [f32; 4] = [0.0, 1.0, 0.0, 1.0];
    pub const BLUE: [f32; 4] = [0.0, 0.0, 1.0, 1.0];
    pub const YELLOW: [f32; 4] = [1.0, 1.0, 0.0, 1.0];
    pub const CYAN: [f32; 4] = [0.0, 1.0, 1.0, 1.0];
    pub const MAGENTA: [f32; 4] = [1.0, 0.0, 1.0, 1.0];

    // Excel-style colors
    pub const EXCEL_BLUE: [f32; 4] = [0.0, 0.32, 0.58, 1.0];
    pub const EXCEL_GREEN: [f32; 4] = [0.0, 0.5, 0.0, 1.0];
    pub const EXCEL_RED: [f32; 4] = [0.75, 0.0, 0.0, 1.0];
    pub const EXCEL_ORANGE: [f32; 4] = [1.0, 0.6, 0.0, 1.0];
    pub const EXCEL_PURPLE: [f32; 4] = [0.5, 0.0, 0.5, 1.0];

    // Grid color
    pub const GRID_GRAY: [f32; 4] = [0.82, 0.82, 0.82, 1.0];
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_hex_to_rgba() {
        assert_eq!(hex_to_rgba("#000000"), Some([0.0, 0.0, 0.0, 1.0]));
        assert_eq!(hex_to_rgba("#ffffff"), Some([1.0, 1.0, 1.0, 1.0]));
        assert_eq!(hex_to_rgba("fff"), Some([1.0, 1.0, 1.0, 1.0]));
    }

    #[test]
    fn test_rgba_to_hex() {
        assert_eq!(rgba_to_hex([0.0, 0.0, 0.0, 1.0]), "#000000");
        assert_eq!(rgba_to_hex([1.0, 1.0, 1.0, 1.0]), "#ffffff");
    }
}
