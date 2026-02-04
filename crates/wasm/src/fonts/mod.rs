//! Embedded fonts for WASM text rendering.
//!
//! This module provides access to embedded fonts for text rendering
//! in the browser without requiring network requests.

use std::collections::HashMap;

/// Font weight variants
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum FontWeight {
    Regular,
    Bold,
}

/// Font style variants
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum FontStyle {
    Normal,
    Italic,
}

/// A font variant key
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct FontVariant {
    pub weight: FontWeight,
    pub style: FontStyle,
}

impl FontVariant {
    pub const REGULAR: Self = Self {
        weight: FontWeight::Regular,
        style: FontStyle::Normal,
    };

    pub const BOLD: Self = Self {
        weight: FontWeight::Bold,
        style: FontStyle::Normal,
    };

    pub const ITALIC: Self = Self {
        weight: FontWeight::Regular,
        style: FontStyle::Italic,
    };

    pub const BOLD_ITALIC: Self = Self {
        weight: FontWeight::Bold,
        style: FontStyle::Italic,
    };
}

/// Font data container
pub struct FontData {
    pub name: &'static str,
    pub data: &'static [u8],
    pub variant: FontVariant,
}

/// Embedded Noto Sans Regular (subset with Latin characters)
///
/// NOTE: In production, this would include actual font bytes.
/// For now, we provide a minimal placeholder that the browser will
/// fallback from, or you can add the actual TTF bytes.
///
/// To add the actual font:
/// 1. Download NotoSans-Regular.ttf from Google Fonts
/// 2. Optionally subset it using pyftsubset or fonttools
/// 3. Include the bytes here with include_bytes!
pub static NOTO_SANS_REGULAR: &[u8] = include_bytes!("NotoSans-Regular.subset.ttf");
pub static NOTO_SANS_BOLD: &[u8] = include_bytes!("NotoSans-Bold.subset.ttf");
pub static NOTO_SANS_ITALIC: &[u8] = include_bytes!("NotoSans-Italic.subset.ttf");
pub static NOTO_SANS_BOLD_ITALIC: &[u8] = include_bytes!("NotoSans-BoldItalic.subset.ttf");

/// Get all embedded fonts
pub fn get_embedded_fonts() -> Vec<FontData> {
    vec![
        FontData {
            name: "Noto Sans",
            data: NOTO_SANS_REGULAR,
            variant: FontVariant::REGULAR,
        },
        FontData {
            name: "Noto Sans",
            data: NOTO_SANS_BOLD,
            variant: FontVariant::BOLD,
        },
        FontData {
            name: "Noto Sans",
            data: NOTO_SANS_ITALIC,
            variant: FontVariant::ITALIC,
        },
        FontData {
            name: "Noto Sans",
            data: NOTO_SANS_BOLD_ITALIC,
            variant: FontVariant::BOLD_ITALIC,
        },
    ]
}

/// Font manager for accessing embedded fonts
pub struct FontManager {
    fonts: HashMap<(String, FontVariant), &'static [u8]>,
    default_font: &'static str,
}

impl Default for FontManager {
    fn default() -> Self {
        Self::new()
    }
}

impl FontManager {
    /// Create a new font manager with embedded fonts
    pub fn new() -> Self {
        let mut fonts = HashMap::new();

        for font in get_embedded_fonts() {
            fonts.insert((font.name.to_string(), font.variant), font.data);
        }

        Self {
            fonts,
            default_font: "Noto Sans",
        }
    }

    /// Get font data for a given family and variant
    pub fn get_font(&self, family: &str, variant: FontVariant) -> Option<&'static [u8]> {
        // First try exact match
        if let Some(data) = self.fonts.get(&(family.to_string(), variant)) {
            return Some(*data);
        }

        // Fallback to default font with same variant
        if let Some(data) = self.fonts.get(&(self.default_font.to_string(), variant)) {
            return Some(*data);
        }

        // Last resort: default font, regular variant
        self.fonts
            .get(&(self.default_font.to_string(), FontVariant::REGULAR))
            .copied()
    }

    /// Get font data for bold/italic styling
    pub fn get_styled_font(&self, family: &str, bold: bool, italic: bool) -> Option<&'static [u8]> {
        let variant = match (bold, italic) {
            (true, true) => FontVariant::BOLD_ITALIC,
            (true, false) => FontVariant::BOLD,
            (false, true) => FontVariant::ITALIC,
            (false, false) => FontVariant::REGULAR,
        };
        self.get_font(family, variant)
    }

    /// Get the default font name
    pub fn default_font(&self) -> &str {
        self.default_font
    }

    /// Check if a font family is available
    pub fn has_font(&self, family: &str) -> bool {
        self.fonts
            .keys()
            .any(|(name, _)| name.eq_ignore_ascii_case(family))
    }

    /// List available font families
    pub fn available_fonts(&self) -> Vec<&str> {
        let mut families: Vec<&str> = self.fonts.keys().map(|(name, _)| name.as_str()).collect();
        families.sort();
        families.dedup();
        families
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_font_variant_constants() {
        assert_eq!(FontVariant::REGULAR.weight, FontWeight::Regular);
        assert_eq!(FontVariant::REGULAR.style, FontStyle::Normal);
        assert_eq!(FontVariant::BOLD.weight, FontWeight::Bold);
        assert_eq!(FontVariant::ITALIC.style, FontStyle::Italic);
    }

    #[test]
    fn test_font_manager_default_font() {
        let manager = FontManager::new();
        assert_eq!(manager.default_font(), "Noto Sans");
    }

    #[test]
    fn test_font_manager_has_font() {
        let manager = FontManager::new();
        assert!(manager.has_font("Noto Sans"));
        assert!(!manager.has_font("Comic Sans"));
    }
}
