//! DOCX document parsing and rendering.
//!
//! Extracts styled content from Word documents for Canvas 2D rendering.
//!
//! This module provides comprehensive DOCX rendering support including:
//! - Full style resolution with inheritance (based on styles, themes, defaults)
//! - Complete run properties (colors, fonts, sizes, super/subscript, etc.)
//! - Complete paragraph properties (alignment, borders, shading, numbering)
//! - Table formatting (cell backgrounds, borders, merging)
//! - Lists and numbering
//! - Images (inline and floating)

use crate::renderer::{points_to_pixels, Color, RenderBackend};
use crate::text_layout::{Paragraph, Rect, TextAlign, TextRun, TextStyle};
use docx_rs::*;
use image::GenericImageView;
use serde::{Deserialize, Serialize};
use std::collections::HashMap;

// ============================================================================
// CONSTANTS AND CONVERSIONS
// ============================================================================

/// Convert twips to points (1 inch = 1440 twips, 1 inch = 72 points)
const TWIPS_TO_POINTS: f32 = 1.0 / 20.0;

/// Convert half-points to points (font sizes in OOXML are in half-points)
const HALF_POINTS_TO_POINTS: f32 = 0.5;

/// Convert EMUs to points (1 inch = 914400 EMU, 1 inch = 72 points)
const EMU_TO_POINTS: f32 = 72.0 / 914400.0;

/// Convert eighths of a point to points (used for some spacing values)
const EIGHTHS_TO_POINTS: f32 = 0.125;

// ============================================================================
// ERROR TYPES
// ============================================================================

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

// ============================================================================
// STYLE RESOLVER - Comprehensive style inheritance and resolution
// ============================================================================

/// Resolved run (character) properties after style inheritance.
#[derive(Clone, Debug, Default)]
pub struct ResolvedRunProps {
    /// Font family (resolved from theme or direct)
    pub font_family: Option<String>,
    /// Font size in points
    pub font_size: Option<f32>,
    /// Bold
    pub bold: Option<bool>,
    /// Italic
    pub italic: Option<bool>,
    /// Underline style
    pub underline: Option<UnderlineStyle>,
    /// Strike-through
    pub strike: Option<bool>,
    /// Double strike-through
    pub double_strike: Option<bool>,
    /// Text color (resolved from theme or direct hex)
    pub color: Option<String>,
    /// Highlight color
    pub highlight: Option<String>,
    /// Shading/background color
    pub shading: Option<String>,
    /// Vertical alignment (superscript/subscript)
    pub vert_align: Option<VerticalAlignType>,
    /// All caps
    pub caps: Option<bool>,
    /// Small caps
    pub small_caps: Option<bool>,
    /// Character spacing in points (positive = expanded, negative = condensed)
    pub character_spacing: Option<f32>,
    /// Text position (raise/lower) in half-points
    pub position: Option<f32>,
    /// Emboss effect
    pub emboss: Option<bool>,
    /// Imprint/engrave effect
    pub imprint: Option<bool>,
    /// Outline effect
    pub outline: Option<bool>,
    /// Shadow effect
    pub shadow: Option<bool>,
    /// Vanish (hidden text)
    pub vanish: Option<bool>,
}

/// Underline style types
#[derive(Clone, Debug, PartialEq)]
pub enum UnderlineStyle {
    None,
    Single,
    Words,
    Double,
    Thick,
    Dotted,
    DottedHeavy,
    Dash,
    DashedHeavy,
    DashLong,
    DashLongHeavy,
    DotDash,
    DashDotHeavy,
    DotDotDash,
    DashDotDotHeavy,
    Wave,
    WavyHeavy,
    WavyDouble,
}

impl Default for UnderlineStyle {
    fn default() -> Self {
        UnderlineStyle::None
    }
}

/// Vertical alignment for superscript/subscript
#[derive(Clone, Debug, PartialEq)]
pub enum VerticalAlignType {
    Baseline,
    Superscript,
    Subscript,
}

impl Default for VerticalAlignType {
    fn default() -> Self {
        VerticalAlignType::Baseline
    }
}

/// Border definition
#[derive(Clone, Debug, Default)]
pub struct BorderDef {
    pub style: BorderStyleType,
    pub width: f32,    // in points
    pub color: String, // hex color
    pub space: f32,    // spacing in points
}

/// Border style types
#[derive(Clone, Debug, PartialEq, Default)]
pub enum BorderStyleType {
    #[default]
    None,
    Single,
    Thick,
    Double,
    Dotted,
    Dashed,
    DotDash,
    DotDotDash,
    Triple,
    ThinThickSmallGap,
    ThickThinSmallGap,
    ThinThickThinSmallGap,
    ThinThickMediumGap,
    ThickThinMediumGap,
    ThinThickThinMediumGap,
    ThinThickLargeGap,
    ThickThinLargeGap,
    ThinThickThinLargeGap,
    Wave,
    DoubleWave,
    DashSmallGap,
    DashDotStroked,
    ThreeDEmboss,
    ThreeDEngrave,
    Outset,
    Inset,
}

/// Resolved paragraph properties after style inheritance.
#[derive(Clone, Debug, Default)]
pub struct ResolvedParaProps {
    /// Text alignment
    pub alignment: Option<AlignmentType>,
    /// Left indent in points
    pub left_indent: Option<f32>,
    /// Right indent in points
    pub right_indent: Option<f32>,
    /// First line indent in points (positive = indent, negative = hanging)
    pub first_line_indent: Option<f32>,
    /// Hanging indent in points
    pub hanging_indent: Option<f32>,
    /// Space before paragraph in points
    pub space_before: Option<f32>,
    /// Space after paragraph in points
    pub space_after: Option<f32>,
    /// Line spacing value
    pub line_spacing: Option<f32>,
    /// Line spacing rule
    pub line_spacing_rule: Option<LineSpacingRule>,
    /// Keep with next paragraph
    pub keep_next: Option<bool>,
    /// Keep lines together
    pub keep_lines: Option<bool>,
    /// Page break before
    pub page_break_before: Option<bool>,
    /// Widow/orphan control
    pub widow_control: Option<bool>,
    /// Paragraph shading/background color
    pub shading: Option<String>,
    /// Paragraph borders
    pub border_top: Option<BorderDef>,
    pub border_bottom: Option<BorderDef>,
    pub border_left: Option<BorderDef>,
    pub border_right: Option<BorderDef>,
    pub border_between: Option<BorderDef>,
    /// Outline level (for headings)
    pub outline_level: Option<u8>,
    /// Style ID (for reference)
    pub style_id: Option<String>,
    /// Numbering properties
    pub numbering_id: Option<i32>,
    pub numbering_level: Option<i32>,
    /// Tab stops
    pub tabs: Vec<TabStop>,
    /// Default run properties for the paragraph
    pub default_run_props: ResolvedRunProps,
}

/// Alignment types
#[derive(Clone, Debug, PartialEq, Default)]
pub enum AlignmentType {
    #[default]
    Left,
    Center,
    Right,
    Both,       // Justify
    Distribute, // Distributed
}

/// Line spacing rules
#[derive(Clone, Debug, PartialEq, Default)]
pub enum LineSpacingRule {
    #[default]
    Auto, // Automatic (value in 240ths of a line)
    Exact,   // Exact (value in twips)
    AtLeast, // At least (value in twips)
}

/// Tab stop definition
#[derive(Clone, Debug)]
pub struct TabStop {
    pub position: f32, // in points
    pub alignment: TabAlignment,
    pub leader: TabLeader,
}

/// Tab alignment
#[derive(Clone, Debug, PartialEq, Default)]
pub enum TabAlignment {
    #[default]
    Left,
    Center,
    Right,
    Decimal,
    Bar,
    Clear,
}

/// Tab leader character
#[derive(Clone, Debug, PartialEq, Default)]
pub enum TabLeader {
    #[default]
    None,
    Dot,
    Hyphen,
    Underscore,
    Heavy,
    MiddleDot,
}

/// Resolved table cell properties
#[derive(Clone, Debug, Default)]
pub struct ResolvedCellProps {
    /// Cell width in points
    pub width: Option<f32>,
    /// Horizontal merge (grid span)
    pub grid_span: Option<u32>,
    /// Vertical merge
    pub v_merge: Option<VMergeType>,
    /// Cell shading/background
    pub shading: Option<String>,
    /// Cell borders
    pub border_top: Option<BorderDef>,
    pub border_bottom: Option<BorderDef>,
    pub border_left: Option<BorderDef>,
    pub border_right: Option<BorderDef>,
    /// Vertical alignment
    pub v_align: Option<CellVerticalAlign>,
    /// Text direction
    pub text_direction: Option<TextDirection>,
    /// Cell margins
    pub margin_top: Option<f32>,
    pub margin_bottom: Option<f32>,
    pub margin_left: Option<f32>,
    pub margin_right: Option<f32>,
    /// No wrap
    pub no_wrap: Option<bool>,
}

/// Vertical merge type
#[derive(Clone, Debug, PartialEq, Default)]
pub enum VMergeType {
    #[default]
    None,
    Restart,  // Start of merged region
    Continue, // Continuation of merged region
}

/// Cell vertical alignment
#[derive(Clone, Debug, PartialEq, Default)]
pub enum CellVerticalAlign {
    #[default]
    Top,
    Center,
    Bottom,
}

/// Text direction in cells
#[derive(Clone, Debug, PartialEq, Default)]
pub enum TextDirection {
    #[default]
    LeftToRightTopToBottom,
    TopToBottomRightToLeft,
    BottomToTopLeftToRight,
}

/// Theme colors as defined in OOXML
#[derive(Clone, Debug)]
pub struct ThemeColors {
    pub dk1: String, // Dark 1 (usually black)
    pub lt1: String, // Light 1 (usually white)
    pub dk2: String, // Dark 2
    pub lt2: String, // Light 2
    pub accent1: String,
    pub accent2: String,
    pub accent3: String,
    pub accent4: String,
    pub accent5: String,
    pub accent6: String,
    pub hlink: String,     // Hyperlink
    pub fol_hlink: String, // Followed hyperlink
}

impl Default for ThemeColors {
    fn default() -> Self {
        // Default Office theme colors
        Self {
            dk1: "000000".to_string(),
            lt1: "FFFFFF".to_string(),
            dk2: "44546A".to_string(),
            lt2: "E7E6E6".to_string(),
            accent1: "4472C4".to_string(), // Blue
            accent2: "ED7D31".to_string(), // Orange
            accent3: "A5A5A5".to_string(), // Gray
            accent4: "FFC000".to_string(), // Gold
            accent5: "5B9BD5".to_string(), // Light blue
            accent6: "70AD47".to_string(), // Green
            hlink: "0563C1".to_string(),
            fol_hlink: "954F72".to_string(),
        }
    }
}

/// Theme fonts
#[derive(Clone, Debug)]
pub struct ThemeFonts {
    pub major_latin: String, // For headings
    pub minor_latin: String, // For body text
    pub major_east_asia: Option<String>,
    pub minor_east_asia: Option<String>,
    pub major_complex: Option<String>,
    pub minor_complex: Option<String>,
}

impl Default for ThemeFonts {
    fn default() -> Self {
        Self {
            major_latin: "Calibri Light".to_string(),
            minor_latin: "Calibri".to_string(),
            major_east_asia: None,
            minor_east_asia: None,
            major_complex: None,
            minor_complex: None,
        }
    }
}

/// Style definition (resolved from docx styles.xml)
#[derive(Clone, Debug, Default)]
pub struct StyleDef {
    pub style_id: String,
    pub name: Option<String>,
    pub style_type: StyleType,
    pub based_on: Option<String>,
    pub next_style: Option<String>,
    pub run_props: ResolvedRunProps,
    pub para_props: ResolvedParaProps,
}

/// Style types
#[derive(Clone, Debug, PartialEq, Default)]
pub enum StyleType {
    #[default]
    Paragraph,
    Character,
    Table,
    Numbering,
}

/// Numbering definition
#[derive(Clone, Debug)]
pub struct NumberingDef {
    pub num_id: i32,
    pub abstract_num_id: i32,
    pub levels: Vec<NumberingLevel>,
}

/// Numbering level definition
#[derive(Clone, Debug)]
pub struct NumberingLevel {
    pub level: i32,
    pub start: i32,
    pub num_fmt: NumberFormat,
    pub lvl_text: String,      // e.g., "%1." or "•"
    pub lvl_jc: AlignmentType, // Justification
    pub indent_left: f32,
    pub hanging: f32,
    pub run_props: ResolvedRunProps,
}

/// Number format types
#[derive(Clone, Debug, PartialEq, Default)]
pub enum NumberFormat {
    #[default]
    Decimal,
    LowerLetter,
    UpperLetter,
    LowerRoman,
    UpperRoman,
    Bullet,
    None,
}

/// The main style resolver that handles all style lookups and inheritance.
pub struct StyleResolver {
    /// Document styles by ID
    styles: HashMap<String, StyleDef>,
    /// Theme colors
    theme_colors: ThemeColors,
    /// Theme fonts
    theme_fonts: ThemeFonts,
    /// Numbering definitions
    numberings: HashMap<i32, NumberingDef>,
    /// Abstract numbering definitions
    abstract_numberings: HashMap<i32, Vec<NumberingLevel>>,
    /// Default paragraph style
    default_para_style: Option<String>,
    /// Default character style
    default_char_style: Option<String>,
    /// Document defaults for run properties
    doc_default_run_props: ResolvedRunProps,
    /// Document defaults for paragraph properties
    doc_default_para_props: ResolvedParaProps,
}

impl StyleResolver {
    /// Create a new StyleResolver from a parsed DOCX document.
    pub fn new(docx: &Docx) -> Self {
        let mut resolver = Self {
            styles: HashMap::new(),
            theme_colors: ThemeColors::default(),
            theme_fonts: ThemeFonts::default(),
            numberings: HashMap::new(),
            abstract_numberings: HashMap::new(),
            default_para_style: None,
            default_char_style: None,
            doc_default_run_props: ResolvedRunProps::default(),
            doc_default_para_props: ResolvedParaProps::default(),
        };

        // Extract theme information
        resolver.extract_themes(docx);

        // Extract document defaults
        resolver.extract_doc_defaults(docx);

        // Extract styles
        resolver.extract_styles(docx);

        // Extract numbering definitions
        resolver.extract_numberings(docx);

        resolver
    }

    /// Extract theme colors and fonts from document.
    fn extract_themes(&mut self, docx: &Docx) {
        // Try to extract theme data via JSON serialization
        if let Ok(json) = serde_json::to_string(&docx.themes) {
            if let Ok(themes) = serde_json::from_str::<serde_json::Value>(&json) {
                if let Some(arr) = themes.as_array() {
                    for theme in arr {
                        // Extract color scheme
                        if let Some(color_scheme) = theme
                            .get("themeElements")
                            .and_then(|te| te.get("colorScheme"))
                        {
                            self.extract_color_scheme(color_scheme);
                        }

                        // Extract font scheme
                        if let Some(font_scheme) = theme
                            .get("themeElements")
                            .and_then(|te| te.get("fontScheme"))
                        {
                            self.extract_font_scheme(font_scheme);
                        }
                    }
                }
            }
        }
    }

    /// Extract color scheme from theme JSON.
    fn extract_color_scheme(&mut self, scheme: &serde_json::Value) {
        if let Some(obj) = scheme.as_object() {
            for (key, value) in obj {
                // Colors can be in different formats - srgbClr or sysClr
                let color = if let Some(srgb) = value.get("srgbClr").and_then(|v| v.as_str()) {
                    srgb.to_string()
                } else if let Some(sys) = value
                    .get("sysClr")
                    .and_then(|v| v.get("lastClr"))
                    .and_then(|v| v.as_str())
                {
                    sys.to_string()
                } else {
                    continue;
                };

                match key.as_str() {
                    "dk1" => self.theme_colors.dk1 = color,
                    "lt1" => self.theme_colors.lt1 = color,
                    "dk2" => self.theme_colors.dk2 = color,
                    "lt2" => self.theme_colors.lt2 = color,
                    "accent1" => self.theme_colors.accent1 = color,
                    "accent2" => self.theme_colors.accent2 = color,
                    "accent3" => self.theme_colors.accent3 = color,
                    "accent4" => self.theme_colors.accent4 = color,
                    "accent5" => self.theme_colors.accent5 = color,
                    "accent6" => self.theme_colors.accent6 = color,
                    "hlink" => self.theme_colors.hlink = color,
                    "folHlink" => self.theme_colors.fol_hlink = color,
                    _ => {}
                }
            }
        }
    }

    /// Extract font scheme from theme JSON.
    fn extract_font_scheme(&mut self, scheme: &serde_json::Value) {
        if let Some(major) = scheme.get("majorFont") {
            if let Some(latin) = major
                .get("latin")
                .and_then(|l| l.get("typeface"))
                .and_then(|t| t.as_str())
            {
                self.theme_fonts.major_latin = latin.to_string();
            }
        }
        if let Some(minor) = scheme.get("minorFont") {
            if let Some(latin) = minor
                .get("latin")
                .and_then(|l| l.get("typeface"))
                .and_then(|t| t.as_str())
            {
                self.theme_fonts.minor_latin = latin.to_string();
            }
        }
    }

    /// Extract document default styles.
    fn extract_doc_defaults(&mut self, docx: &Docx) {
        // Document defaults are in docx.styles.doc_defaults
        if let Ok(json) = serde_json::to_string(&docx.styles.doc_defaults) {
            if let Ok(defaults) = serde_json::from_str::<serde_json::Value>(&json) {
                // Run properties defaults
                if let Some(run_props) = defaults
                    .get("runPropertyDefault")
                    .and_then(|rpd| rpd.get("runProperty"))
                {
                    self.doc_default_run_props = self.parse_run_props_json(run_props);
                }

                // Paragraph properties defaults
                if let Some(para_props) = defaults
                    .get("paragraphPropertyDefault")
                    .and_then(|ppd| ppd.get("paragraphProperty"))
                {
                    self.doc_default_para_props = self.parse_para_props_json(para_props);
                }
            }
        }
    }

    /// Extract style definitions from document.
    fn extract_styles(&mut self, docx: &Docx) {
        if let Ok(json) = serde_json::to_string(&docx.styles.styles) {
            if let Ok(styles) = serde_json::from_str::<serde_json::Value>(&json) {
                if let Some(arr) = styles.as_array() {
                    for style in arr {
                        if let Some(style_def) = self.parse_style_json(style) {
                            // Check for default styles
                            if style
                                .get("default")
                                .and_then(|d| d.as_bool())
                                .unwrap_or(false)
                            {
                                match style_def.style_type {
                                    StyleType::Paragraph => {
                                        self.default_para_style = Some(style_def.style_id.clone());
                                    }
                                    StyleType::Character => {
                                        self.default_char_style = Some(style_def.style_id.clone());
                                    }
                                    _ => {}
                                }
                            }

                            self.styles.insert(style_def.style_id.clone(), style_def);
                        }
                    }
                }
            }
        }
    }

    /// Parse a style from JSON.
    fn parse_style_json(&self, style: &serde_json::Value) -> Option<StyleDef> {
        let style_id = style.get("styleId").and_then(|s| s.as_str())?.to_string();
        let name = style
            .get("name")
            .and_then(|n| n.get("val"))
            .and_then(|v| v.as_str())
            .map(|s| s.to_string());
        let based_on = style
            .get("basedOn")
            .and_then(|b| b.get("val"))
            .and_then(|v| v.as_str())
            .map(|s| s.to_string());
        let next_style = style
            .get("next")
            .and_then(|n| n.get("val"))
            .and_then(|v| v.as_str())
            .map(|s| s.to_string());

        let style_type = match style.get("styleType").and_then(|t| t.as_str()) {
            Some("paragraph") => StyleType::Paragraph,
            Some("character") => StyleType::Character,
            Some("table") => StyleType::Table,
            Some("numbering") => StyleType::Numbering,
            _ => StyleType::Paragraph,
        };

        let run_props = style
            .get("runProperty")
            .map(|rp| self.parse_run_props_json(rp))
            .unwrap_or_default();

        let para_props = style
            .get("paragraphProperty")
            .map(|pp| self.parse_para_props_json(pp))
            .unwrap_or_default();

        Some(StyleDef {
            style_id,
            name,
            style_type,
            based_on,
            next_style,
            run_props,
            para_props,
        })
    }

    /// Parse run properties from JSON.
    fn parse_run_props_json(&self, props: &serde_json::Value) -> ResolvedRunProps {
        let mut resolved = ResolvedRunProps::default();

        // Font size (sz is in half-points)
        if let Some(sz) = props.get("sz").and_then(|s| s.as_f64()) {
            resolved.font_size = Some(sz as f32 * HALF_POINTS_TO_POINTS);
        }

        // Bold
        if props.get("bold").is_some() {
            resolved.bold = Some(true);
        }

        // Italic
        if props.get("italic").is_some() {
            resolved.italic = Some(true);
        }

        // Underline
        if let Some(underline) = props.get("underline") {
            resolved.underline = Some(self.parse_underline_json(underline));
        }

        // Strike-through
        if props.get("strike").is_some() {
            resolved.strike = Some(true);
        }

        // Double strike-through
        if props.get("dstrike").is_some() {
            resolved.double_strike = Some(true);
        }

        // Color
        if let Some(color) = props.get("color") {
            resolved.color = self.parse_color_json(color);
        }

        // Highlight
        if let Some(highlight) = props.get("highlight") {
            if let Some(val) = highlight.as_str() {
                resolved.highlight = Some(self.highlight_name_to_hex(val));
            }
        }

        // Shading
        if let Some(shading) = props.get("shading") {
            if let Some(fill) = shading.get("fill").and_then(|f| f.as_str()) {
                if fill != "auto" {
                    resolved.shading = Some(format!("#{}", fill));
                }
            }
        }

        // Vertical alignment (superscript/subscript)
        if let Some(vert_align) = props
            .get("vertAlign")
            .and_then(|v| v.get("val"))
            .and_then(|v| v.as_str())
        {
            resolved.vert_align = Some(match vert_align {
                "superscript" => VerticalAlignType::Superscript,
                "subscript" => VerticalAlignType::Subscript,
                _ => VerticalAlignType::Baseline,
            });
        }

        // Caps
        if props.get("caps").is_some() {
            resolved.caps = Some(true);
        }

        // Small caps
        if props.get("smallCaps").is_some() {
            resolved.small_caps = Some(true);
        }

        // Character spacing (spacing is in twips)
        if let Some(spacing) = props
            .get("spacing")
            .and_then(|s| s.get("val"))
            .and_then(|v| v.as_i64())
        {
            resolved.character_spacing = Some(spacing as f32 * TWIPS_TO_POINTS);
        }

        // Fonts
        if let Some(fonts) = props.get("fonts") {
            resolved.font_family = self.parse_fonts_json(fonts);
        }

        // Vanish (hidden)
        if props.get("vanish").is_some() {
            resolved.vanish = Some(true);
        }

        // Shadow
        if props.get("shadow").is_some() {
            resolved.shadow = Some(true);
        }

        // Outline
        if props.get("outline").is_some() {
            resolved.outline = Some(true);
        }

        // Emboss
        if props.get("emboss").is_some() {
            resolved.emboss = Some(true);
        }

        // Imprint
        if props.get("imprint").is_some() {
            resolved.imprint = Some(true);
        }

        resolved
    }

    /// Parse underline from JSON.
    fn parse_underline_json(&self, underline: &serde_json::Value) -> UnderlineStyle {
        let val = underline
            .get("val")
            .and_then(|v| v.as_str())
            .unwrap_or("single");
        match val {
            "none" => UnderlineStyle::None,
            "single" => UnderlineStyle::Single,
            "words" => UnderlineStyle::Words,
            "double" => UnderlineStyle::Double,
            "thick" => UnderlineStyle::Thick,
            "dotted" => UnderlineStyle::Dotted,
            "dottedHeavy" => UnderlineStyle::DottedHeavy,
            "dash" => UnderlineStyle::Dash,
            "dashedHeavy" => UnderlineStyle::DashedHeavy,
            "dashLong" => UnderlineStyle::DashLong,
            "dashLongHeavy" => UnderlineStyle::DashLongHeavy,
            "dotDash" => UnderlineStyle::DotDash,
            "dashDotHeavy" => UnderlineStyle::DashDotHeavy,
            "dotDotDash" => UnderlineStyle::DotDotDash,
            "dashDotDotHeavy" => UnderlineStyle::DashDotDotHeavy,
            "wave" => UnderlineStyle::Wave,
            "wavyHeavy" => UnderlineStyle::WavyHeavy,
            "wavyDouble" => UnderlineStyle::WavyDouble,
            _ => UnderlineStyle::Single,
        }
    }

    /// Parse color from JSON, resolving theme colors.
    fn parse_color_json(&self, color: &serde_json::Value) -> Option<String> {
        // Check for direct hex color
        if let Some(val) = color.get("val").and_then(|v| v.as_str()) {
            if val != "auto" {
                return Some(format!("#{}", val));
            }
        }

        // Check for theme color
        if let Some(theme_color) = color.get("themeColor").and_then(|t| t.as_str()) {
            let base_color = self.resolve_theme_color(theme_color);

            // Apply tint/shade modifiers if present
            let tint = color
                .get("themeTint")
                .and_then(|t| t.as_str())
                .and_then(|t| u8::from_str_radix(t, 16).ok());
            let shade = color
                .get("themeShade")
                .and_then(|s| s.as_str())
                .and_then(|s| u8::from_str_radix(s, 16).ok());

            if let Some(tint_val) = tint {
                return Some(self.apply_tint(&base_color, tint_val));
            }
            if let Some(shade_val) = shade {
                return Some(self.apply_shade(&base_color, shade_val));
            }

            return Some(format!("#{}", base_color));
        }

        None
    }

    /// Resolve a theme color name to hex.
    pub fn resolve_theme_color(&self, theme_color: &str) -> String {
        match theme_color {
            "dark1" | "dk1" => self.theme_colors.dk1.clone(),
            "light1" | "lt1" => self.theme_colors.lt1.clone(),
            "dark2" | "dk2" => self.theme_colors.dk2.clone(),
            "light2" | "lt2" => self.theme_colors.lt2.clone(),
            "accent1" => self.theme_colors.accent1.clone(),
            "accent2" => self.theme_colors.accent2.clone(),
            "accent3" => self.theme_colors.accent3.clone(),
            "accent4" => self.theme_colors.accent4.clone(),
            "accent5" => self.theme_colors.accent5.clone(),
            "accent6" => self.theme_colors.accent6.clone(),
            "hyperlink" | "hlink" => self.theme_colors.hlink.clone(),
            "followedHyperlink" | "folHlink" => self.theme_colors.fol_hlink.clone(),
            _ => "000000".to_string(),
        }
    }

    /// Apply tint to a color (lighten).
    fn apply_tint(&self, hex_color: &str, tint: u8) -> String {
        let tint_factor = tint as f32 / 255.0;
        self.modify_color(hex_color, |c| c + (1.0 - c) * tint_factor)
    }

    /// Apply shade to a color (darken).
    fn apply_shade(&self, hex_color: &str, shade: u8) -> String {
        let shade_factor = shade as f32 / 255.0;
        self.modify_color(hex_color, |c| c * shade_factor)
    }

    /// Modify RGB components of a color.
    fn modify_color<F>(&self, hex_color: &str, modifier: F) -> String
    where
        F: Fn(f32) -> f32,
    {
        let hex = hex_color.trim_start_matches('#');
        if hex.len() < 6 {
            return format!("#{}", hex_color);
        }

        let r = u8::from_str_radix(&hex[0..2], 16).unwrap_or(0) as f32 / 255.0;
        let g = u8::from_str_radix(&hex[2..4], 16).unwrap_or(0) as f32 / 255.0;
        let b = u8::from_str_radix(&hex[4..6], 16).unwrap_or(0) as f32 / 255.0;

        let r_new = (modifier(r) * 255.0).clamp(0.0, 255.0) as u8;
        let g_new = (modifier(g) * 255.0).clamp(0.0, 255.0) as u8;
        let b_new = (modifier(b) * 255.0).clamp(0.0, 255.0) as u8;

        format!("#{:02X}{:02X}{:02X}", r_new, g_new, b_new)
    }

    /// Parse fonts from JSON, resolving theme fonts.
    fn parse_fonts_json(&self, fonts: &serde_json::Value) -> Option<String> {
        // Try ascii first, then hAnsi, then eastAsia
        if let Some(ascii) = fonts.get("ascii").and_then(|a| a.as_str()) {
            // Check for theme font references
            return Some(self.resolve_font_name(ascii));
        }
        if let Some(h_ansi) = fonts.get("hAnsi").and_then(|h| h.as_str()) {
            return Some(self.resolve_font_name(h_ansi));
        }
        if let Some(ascii_theme) = fonts.get("asciiTheme").and_then(|a| a.as_str()) {
            return Some(self.resolve_theme_font(ascii_theme));
        }
        None
    }

    /// Resolve a font name, handling theme references.
    fn resolve_font_name(&self, name: &str) -> String {
        match name {
            "majorHAnsi" | "majorAscii" => self.theme_fonts.major_latin.clone(),
            "minorHAnsi" | "minorAscii" => self.theme_fonts.minor_latin.clone(),
            _ => name.to_string(),
        }
    }

    /// Resolve a theme font name.
    fn resolve_theme_font(&self, theme_font: &str) -> String {
        match theme_font {
            "majorHAnsi" | "majorAscii" | "majorBidi" => self.theme_fonts.major_latin.clone(),
            "minorHAnsi" | "minorAscii" | "minorBidi" => self.theme_fonts.minor_latin.clone(),
            _ => self.theme_fonts.minor_latin.clone(),
        }
    }

    /// Parse paragraph properties from JSON.
    fn parse_para_props_json(&self, props: &serde_json::Value) -> ResolvedParaProps {
        let mut resolved = ResolvedParaProps::default();

        // Alignment
        if let Some(jc) = props.get("alignment").and_then(|a| a.as_str()) {
            resolved.alignment = Some(match jc {
                "left" | "start" => AlignmentType::Left,
                "center" => AlignmentType::Center,
                "right" | "end" => AlignmentType::Right,
                "both" | "justify" => AlignmentType::Both,
                "distribute" => AlignmentType::Distribute,
                _ => AlignmentType::Left,
            });
        }

        // Indentation
        if let Some(indent) = props.get("indent") {
            if let Some(left) = indent.get("start").and_then(|s| s.as_i64()) {
                resolved.left_indent = Some(left as f32 * TWIPS_TO_POINTS);
            }
            if let Some(right) = indent.get("end").and_then(|e| e.as_i64()) {
                resolved.right_indent = Some(right as f32 * TWIPS_TO_POINTS);
            }
            if let Some(first) = indent.get("firstLineChars").and_then(|f| f.as_i64()) {
                // First line in hundredths of a character width
                resolved.first_line_indent = Some(first as f32 / 100.0 * 11.0); // Approximate
            }
            if let Some(first) = indent.get("firstLine").and_then(|f| f.as_i64()) {
                resolved.first_line_indent = Some(first as f32 * TWIPS_TO_POINTS);
            }
            if let Some(hanging) = indent.get("hanging").and_then(|h| h.as_i64()) {
                resolved.hanging_indent = Some(hanging as f32 * TWIPS_TO_POINTS);
            }
        }

        // Spacing
        if let Some(spacing) = props.get("lineSpacing") {
            if let Some(before) = spacing.get("before").and_then(|b| b.as_i64()) {
                resolved.space_before = Some(before as f32 * TWIPS_TO_POINTS);
            }
            if let Some(after) = spacing.get("after").and_then(|a| a.as_i64()) {
                resolved.space_after = Some(after as f32 * TWIPS_TO_POINTS);
            }
            if let Some(line) = spacing.get("line").and_then(|l| l.as_i64()) {
                let rule = spacing
                    .get("lineRule")
                    .and_then(|r| r.as_str())
                    .unwrap_or("auto");
                resolved.line_spacing_rule = Some(match rule {
                    "exact" => LineSpacingRule::Exact,
                    "atLeast" => LineSpacingRule::AtLeast,
                    _ => LineSpacingRule::Auto,
                });
                // Convert based on rule
                match resolved.line_spacing_rule {
                    Some(LineSpacingRule::Auto) => {
                        // Auto is in 240ths of a line
                        resolved.line_spacing = Some(line as f32 / 240.0);
                    }
                    _ => {
                        // Exact/AtLeast are in twips
                        resolved.line_spacing = Some(line as f32 * TWIPS_TO_POINTS);
                    }
                }
            }
        }

        // Shading
        if let Some(shading) = props.get("shading") {
            if let Some(fill) = shading.get("fill").and_then(|f| f.as_str()) {
                if fill != "auto" {
                    resolved.shading = Some(format!("#{}", fill));
                }
            }
        }

        // Borders
        if let Some(borders) = props.get("borders") {
            resolved.border_top = borders.get("top").map(|b| self.parse_border_json(b));
            resolved.border_bottom = borders.get("bottom").map(|b| self.parse_border_json(b));
            resolved.border_left = borders.get("left").map(|b| self.parse_border_json(b));
            resolved.border_right = borders.get("right").map(|b| self.parse_border_json(b));
            resolved.border_between = borders.get("between").map(|b| self.parse_border_json(b));
        }

        // Style reference
        if let Some(style) = props
            .get("style")
            .and_then(|s| s.get("val"))
            .and_then(|v| v.as_str())
        {
            resolved.style_id = Some(style.to_string());
        }

        // Numbering
        if let Some(num_pr) = props.get("numberingProperty") {
            if let Some(num_id) = num_pr.get("id").and_then(|i| i.as_i64()) {
                resolved.numbering_id = Some(num_id as i32);
            }
            if let Some(ilvl) = num_pr.get("level").and_then(|l| l.as_i64()) {
                resolved.numbering_level = Some(ilvl as i32);
            }
        }

        // Keep with next
        if props.get("keepNext").is_some() {
            resolved.keep_next = Some(true);
        }

        // Keep lines
        if props.get("keepLines").is_some() {
            resolved.keep_lines = Some(true);
        }

        // Page break before
        if props.get("pageBreakBefore").is_some() {
            resolved.page_break_before = Some(true);
        }

        // Widow control
        if props.get("widowControl").is_some() {
            resolved.widow_control = Some(true);
        }

        // Outline level
        if let Some(outline_lvl) = props
            .get("outlineLvl")
            .and_then(|o| o.get("val"))
            .and_then(|v| v.as_i64())
        {
            resolved.outline_level = Some(outline_lvl as u8);
        }

        resolved
    }

    /// Parse border from JSON.
    fn parse_border_json(&self, border: &serde_json::Value) -> BorderDef {
        let mut def = BorderDef::default();

        if let Some(val) = border.get("val").and_then(|v| v.as_str()) {
            def.style = self.parse_border_style(val);
        }

        if let Some(sz) = border.get("sz").and_then(|s| s.as_i64()) {
            // Border size is in eighths of a point
            def.width = sz as f32 * EIGHTHS_TO_POINTS;
        }

        if let Some(color) = border.get("color").and_then(|c| c.as_str()) {
            if color != "auto" {
                def.color = format!("#{}", color);
            } else {
                def.color = "#000000".to_string();
            }
        }

        if let Some(space) = border.get("space").and_then(|s| s.as_i64()) {
            def.space = space as f32; // in points
        }

        def
    }

    /// Parse border style string.
    fn parse_border_style(&self, style: &str) -> BorderStyleType {
        match style {
            "nil" | "none" => BorderStyleType::None,
            "single" => BorderStyleType::Single,
            "thick" => BorderStyleType::Thick,
            "double" => BorderStyleType::Double,
            "dotted" => BorderStyleType::Dotted,
            "dashed" => BorderStyleType::Dashed,
            "dotDash" => BorderStyleType::DotDash,
            "dotDotDash" => BorderStyleType::DotDotDash,
            "triple" => BorderStyleType::Triple,
            "wave" => BorderStyleType::Wave,
            "doubleWave" => BorderStyleType::DoubleWave,
            "dashSmallGap" => BorderStyleType::DashSmallGap,
            "threeDEmboss" => BorderStyleType::ThreeDEmboss,
            "threeDEngrave" => BorderStyleType::ThreeDEngrave,
            "outset" => BorderStyleType::Outset,
            "inset" => BorderStyleType::Inset,
            _ => BorderStyleType::Single,
        }
    }

    /// Extract numbering definitions.
    fn extract_numberings(&mut self, docx: &Docx) {
        // Parse abstract numbering definitions
        if let Ok(json) = serde_json::to_string(&docx.numberings.abstract_nums) {
            if let Ok(abstract_nums) = serde_json::from_str::<serde_json::Value>(&json) {
                if let Some(arr) = abstract_nums.as_array() {
                    for abs_num in arr {
                        if let Some(id) = abs_num.get("id").and_then(|i| i.as_i64()) {
                            let levels = self.parse_numbering_levels(abs_num);
                            self.abstract_numberings.insert(id as i32, levels);
                        }
                    }
                }
            }
        }

        // Parse numbering instances
        if let Ok(json) = serde_json::to_string(&docx.numberings.numberings) {
            if let Ok(nums) = serde_json::from_str::<serde_json::Value>(&json) {
                if let Some(arr) = nums.as_array() {
                    for num in arr {
                        if let (Some(num_id), Some(abstract_num_id)) = (
                            num.get("id").and_then(|i| i.as_i64()),
                            num.get("abstractNumId").and_then(|a| a.as_i64()),
                        ) {
                            let levels = self
                                .abstract_numberings
                                .get(&(abstract_num_id as i32))
                                .cloned()
                                .unwrap_or_default();

                            self.numberings.insert(
                                num_id as i32,
                                NumberingDef {
                                    num_id: num_id as i32,
                                    abstract_num_id: abstract_num_id as i32,
                                    levels,
                                },
                            );
                        }
                    }
                }
            }
        }
    }

    /// Parse numbering levels from JSON.
    fn parse_numbering_levels(&self, abs_num: &serde_json::Value) -> Vec<NumberingLevel> {
        let mut levels = Vec::new();

        if let Some(lvls) = abs_num.get("levels").and_then(|l| l.as_array()) {
            for lvl in lvls {
                if let Some(level) = self.parse_numbering_level(lvl) {
                    levels.push(level);
                }
            }
        }

        levels
    }

    /// Parse a single numbering level.
    fn parse_numbering_level(&self, lvl: &serde_json::Value) -> Option<NumberingLevel> {
        let level = lvl.get("level").and_then(|l| l.as_i64())? as i32;

        let start = lvl.get("start").and_then(|s| s.as_i64()).unwrap_or(1) as i32;

        let num_fmt = match lvl
            .get("numFmt")
            .and_then(|n| n.get("val"))
            .and_then(|v| v.as_str())
        {
            Some("decimal") => NumberFormat::Decimal,
            Some("lowerLetter") => NumberFormat::LowerLetter,
            Some("upperLetter") => NumberFormat::UpperLetter,
            Some("lowerRoman") => NumberFormat::LowerRoman,
            Some("upperRoman") => NumberFormat::UpperRoman,
            Some("bullet") => NumberFormat::Bullet,
            Some("none") => NumberFormat::None,
            _ => NumberFormat::Decimal,
        };

        let lvl_text = lvl
            .get("levelText")
            .and_then(|t| t.get("val"))
            .and_then(|v| v.as_str())
            .unwrap_or("%1.")
            .to_string();

        let lvl_jc = match lvl
            .get("levelJc")
            .and_then(|j| j.get("val"))
            .and_then(|v| v.as_str())
        {
            Some("left") => AlignmentType::Left,
            Some("center") => AlignmentType::Center,
            Some("right") => AlignmentType::Right,
            _ => AlignmentType::Left,
        };

        let mut indent_left = 0.0;
        let mut hanging = 0.0;

        if let Some(indent) = lvl.get("paragraphProperty").and_then(|pp| pp.get("indent")) {
            if let Some(left) = indent.get("left").and_then(|l| l.as_i64()) {
                indent_left = left as f32 * TWIPS_TO_POINTS;
            }
            if let Some(hang) = indent.get("hanging").and_then(|h| h.as_i64()) {
                hanging = hang as f32 * TWIPS_TO_POINTS;
            }
        }

        let run_props = lvl
            .get("runProperty")
            .map(|rp| self.parse_run_props_json(rp))
            .unwrap_or_default();

        Some(NumberingLevel {
            level,
            start,
            num_fmt,
            lvl_text,
            lvl_jc,
            indent_left,
            hanging,
            run_props,
        })
    }

    /// Convert highlight color name to hex.
    fn highlight_name_to_hex(&self, name: &str) -> String {
        match name.to_lowercase().as_str() {
            "yellow" => "#FFFF00".to_string(),
            "green" => "#00FF00".to_string(),
            "cyan" => "#00FFFF".to_string(),
            "magenta" => "#FF00FF".to_string(),
            "blue" => "#0000FF".to_string(),
            "red" => "#FF0000".to_string(),
            "darkblue" | "darkBlue" => "#000080".to_string(),
            "darkcyan" | "darkCyan" => "#008080".to_string(),
            "darkgreen" | "darkGreen" => "#008000".to_string(),
            "darkmagenta" | "darkMagenta" => "#800080".to_string(),
            "darkred" | "darkRed" => "#800000".to_string(),
            "darkyellow" | "darkYellow" => "#808000".to_string(),
            "darkgray" | "darkGray" => "#808080".to_string(),
            "lightgray" | "lightGray" => "#C0C0C0".to_string(),
            "black" => "#000000".to_string(),
            "white" => "#FFFFFF".to_string(),
            _ => "#FFFF00".to_string(),
        }
    }

    /// Resolve complete run properties for a given style and direct properties.
    ///
    /// Resolution order (lowest to highest priority):
    /// 1. Document defaults
    /// 2. Paragraph style's run properties
    /// 3. Character style (if any)
    /// 4. Direct run properties
    pub fn resolve_run_props(
        &self,
        para_style_id: Option<&str>,
        char_style_id: Option<&str>,
        direct_props: &ResolvedRunProps,
    ) -> ResolvedRunProps {
        let mut resolved = self.doc_default_run_props.clone();

        // Apply paragraph style's run properties
        if let Some(style_id) = para_style_id {
            let style_props = self.get_style_run_props(style_id);
            resolved = Self::merge_run_props(&resolved, &style_props);
        }

        // Apply character style
        if let Some(style_id) = char_style_id {
            let style_props = self.get_style_run_props(style_id);
            resolved = Self::merge_run_props(&resolved, &style_props);
        }

        // Apply direct properties (highest priority)
        resolved = Self::merge_run_props(&resolved, direct_props);

        resolved
    }

    /// Get run properties for a style, including inherited properties.
    fn get_style_run_props(&self, style_id: &str) -> ResolvedRunProps {
        let mut chain = Vec::new();
        let mut current_id = Some(style_id.to_string());

        // Build inheritance chain
        while let Some(id) = current_id {
            if chain.contains(&id) {
                break; // Prevent infinite loops
            }
            chain.push(id.clone());
            current_id = self.styles.get(&id).and_then(|s| s.based_on.clone());
        }

        // Apply from base to derived
        let mut resolved = ResolvedRunProps::default();
        for id in chain.into_iter().rev() {
            if let Some(style) = self.styles.get(&id) {
                resolved = Self::merge_run_props(&resolved, &style.run_props);
            }
        }

        resolved
    }

    /// Merge two run property sets (source overrides base where defined).
    fn merge_run_props(base: &ResolvedRunProps, source: &ResolvedRunProps) -> ResolvedRunProps {
        ResolvedRunProps {
            font_family: source
                .font_family
                .clone()
                .or_else(|| base.font_family.clone()),
            font_size: source.font_size.or(base.font_size),
            bold: source.bold.or(base.bold),
            italic: source.italic.or(base.italic),
            underline: source.underline.clone().or_else(|| base.underline.clone()),
            strike: source.strike.or(base.strike),
            double_strike: source.double_strike.or(base.double_strike),
            color: source.color.clone().or_else(|| base.color.clone()),
            highlight: source.highlight.clone().or_else(|| base.highlight.clone()),
            shading: source.shading.clone().or_else(|| base.shading.clone()),
            vert_align: source
                .vert_align
                .clone()
                .or_else(|| base.vert_align.clone()),
            caps: source.caps.or(base.caps),
            small_caps: source.small_caps.or(base.small_caps),
            character_spacing: source.character_spacing.or(base.character_spacing),
            position: source.position.or(base.position),
            emboss: source.emboss.or(base.emboss),
            imprint: source.imprint.or(base.imprint),
            outline: source.outline.or(base.outline),
            shadow: source.shadow.or(base.shadow),
            vanish: source.vanish.or(base.vanish),
        }
    }

    /// Resolve complete paragraph properties.
    pub fn resolve_para_props(
        &self,
        style_id: Option<&str>,
        direct_props: &ResolvedParaProps,
    ) -> ResolvedParaProps {
        let mut resolved = self.doc_default_para_props.clone();

        // Apply style properties (with inheritance)
        if let Some(style_id) = style_id {
            let style_props = self.get_style_para_props(style_id);
            resolved = Self::merge_para_props(&resolved, &style_props);
        }

        // Apply direct properties
        resolved = Self::merge_para_props(&resolved, direct_props);

        resolved
    }

    /// Get paragraph properties for a style, including inherited properties.
    fn get_style_para_props(&self, style_id: &str) -> ResolvedParaProps {
        let mut chain = Vec::new();
        let mut current_id = Some(style_id.to_string());

        while let Some(id) = current_id {
            if chain.contains(&id) {
                break;
            }
            chain.push(id.clone());
            current_id = self.styles.get(&id).and_then(|s| s.based_on.clone());
        }

        let mut resolved = ResolvedParaProps::default();
        for id in chain.into_iter().rev() {
            if let Some(style) = self.styles.get(&id) {
                resolved = Self::merge_para_props(&resolved, &style.para_props);
            }
        }

        resolved
    }

    /// Merge paragraph properties.
    fn merge_para_props(base: &ResolvedParaProps, source: &ResolvedParaProps) -> ResolvedParaProps {
        ResolvedParaProps {
            alignment: source.alignment.clone().or_else(|| base.alignment.clone()),
            left_indent: source.left_indent.or(base.left_indent),
            right_indent: source.right_indent.or(base.right_indent),
            first_line_indent: source.first_line_indent.or(base.first_line_indent),
            hanging_indent: source.hanging_indent.or(base.hanging_indent),
            space_before: source.space_before.or(base.space_before),
            space_after: source.space_after.or(base.space_after),
            line_spacing: source.line_spacing.or(base.line_spacing),
            line_spacing_rule: source
                .line_spacing_rule
                .clone()
                .or_else(|| base.line_spacing_rule.clone()),
            keep_next: source.keep_next.or(base.keep_next),
            keep_lines: source.keep_lines.or(base.keep_lines),
            page_break_before: source.page_break_before.or(base.page_break_before),
            widow_control: source.widow_control.or(base.widow_control),
            shading: source.shading.clone().or_else(|| base.shading.clone()),
            border_top: source
                .border_top
                .clone()
                .or_else(|| base.border_top.clone()),
            border_bottom: source
                .border_bottom
                .clone()
                .or_else(|| base.border_bottom.clone()),
            border_left: source
                .border_left
                .clone()
                .or_else(|| base.border_left.clone()),
            border_right: source
                .border_right
                .clone()
                .or_else(|| base.border_right.clone()),
            border_between: source
                .border_between
                .clone()
                .or_else(|| base.border_between.clone()),
            outline_level: source.outline_level.or(base.outline_level),
            style_id: source.style_id.clone().or_else(|| base.style_id.clone()),
            numbering_id: source.numbering_id.or(base.numbering_id),
            numbering_level: source.numbering_level.or(base.numbering_level),
            tabs: if source.tabs.is_empty() {
                base.tabs.clone()
            } else {
                source.tabs.clone()
            },
            default_run_props: Self::merge_run_props(
                &base.default_run_props,
                &source.default_run_props,
            ),
        }
    }

    /// Get numbering definition by ID.
    pub fn get_numbering(&self, num_id: i32) -> Option<&NumberingDef> {
        self.numberings.get(&num_id)
    }

    /// Format a number according to the numbering format.
    pub fn format_number(&self, value: i32, format: &NumberFormat) -> String {
        match format {
            NumberFormat::Decimal => value.to_string(),
            NumberFormat::LowerLetter => self.number_to_letter(value, false),
            NumberFormat::UpperLetter => self.number_to_letter(value, true),
            NumberFormat::LowerRoman => self.number_to_roman(value, false),
            NumberFormat::UpperRoman => self.number_to_roman(value, true),
            NumberFormat::Bullet => "•".to_string(),
            NumberFormat::None => String::new(),
        }
    }

    /// Convert number to letter (1=a, 2=b, ..., 27=aa, etc.)
    fn number_to_letter(&self, value: i32, uppercase: bool) -> String {
        let mut result = String::new();
        let mut n = value;

        while n > 0 {
            n -= 1;
            let letter = ((n % 26) as u8 + if uppercase { b'A' } else { b'a' }) as char;
            result.insert(0, letter);
            n /= 26;
        }

        result
    }

    /// Convert number to Roman numerals.
    fn number_to_roman(&self, value: i32, uppercase: bool) -> String {
        let numerals = [
            (1000, "m"),
            (900, "cm"),
            (500, "d"),
            (400, "cd"),
            (100, "c"),
            (90, "xc"),
            (50, "l"),
            (40, "xl"),
            (10, "x"),
            (9, "ix"),
            (5, "v"),
            (4, "iv"),
            (1, "i"),
        ];

        let mut result = String::new();
        let mut n = value;

        for (val, numeral) in &numerals {
            while n >= *val {
                result.push_str(numeral);
                n -= val;
            }
        }

        if uppercase {
            result.to_uppercase()
        } else {
            result
        }
    }

    /// Get the theme colors.
    pub fn theme_colors(&self) -> &ThemeColors {
        &self.theme_colors
    }

    /// Get the theme fonts.
    pub fn theme_fonts(&self) -> &ThemeFonts {
        &self.theme_fonts
    }
}

// ============================================================================
// STYLED DOCUMENT STRUCTURES
// ============================================================================

/// A styled text run extracted from DOCX with comprehensive formatting.
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct StyledRun {
    /// The text content
    pub text: String,
    /// Font family name (resolved from theme if necessary)
    pub font_family: String,
    /// Font size in points
    pub font_size: f32,
    /// Bold formatting
    pub bold: bool,
    /// Italic formatting
    pub italic: bool,
    /// Underline style
    pub underline: bool,
    /// Underline style type (for rendering different underline styles)
    #[serde(default)]
    pub underline_style: String,
    /// Strike-through
    pub strikethrough: bool,
    /// Double strike-through
    #[serde(default)]
    pub double_strikethrough: bool,
    /// Text color (hex, e.g., "#FF0000")
    pub color: String,
    /// Highlight/background color (hex)
    pub highlight: Option<String>,
    /// Shading/background color (from shading element)
    #[serde(default)]
    pub shading: Option<String>,
    /// Vertical alignment: "baseline", "superscript", or "subscript"
    #[serde(default)]
    pub vertical_align: String,
    /// All caps
    #[serde(default)]
    pub all_caps: bool,
    /// Small caps
    #[serde(default)]
    pub small_caps: bool,
    /// Character spacing adjustment in points (positive = expanded)
    #[serde(default)]
    pub character_spacing: f32,
    /// Text position offset in points (positive = raised)
    #[serde(default)]
    pub position: f32,
    /// Hidden text
    #[serde(default)]
    pub hidden: bool,
    /// Emboss effect
    #[serde(default)]
    pub emboss: bool,
    /// Imprint/engrave effect
    #[serde(default)]
    pub imprint: bool,
    /// Outline effect
    #[serde(default)]
    pub outline: bool,
    /// Shadow effect
    #[serde(default)]
    pub shadow: bool,
    /// Is this run a hyperlink?
    #[serde(default)]
    pub is_hyperlink: bool,
    /// Hyperlink target URL
    #[serde(default)]
    pub hyperlink_url: Option<String>,
    /// Is this run a bookmark?
    #[serde(default)]
    pub bookmark_name: Option<String>,
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
            underline_style: "single".to_string(),
            strikethrough: false,
            double_strikethrough: false,
            color: "#000000".to_string(),
            highlight: None,
            shading: None,
            vertical_align: "baseline".to_string(),
            all_caps: false,
            small_caps: false,
            character_spacing: 0.0,
            position: 0.0,
            hidden: false,
            emboss: false,
            imprint: false,
            outline: false,
            shadow: false,
            is_hyperlink: false,
            hyperlink_url: None,
            bookmark_name: None,
        }
    }
}

impl StyledRun {
    /// Convert to TextStyle for rendering.
    pub fn to_text_style(&self) -> TextStyle {
        let mut style = TextStyle::new(&self.font_family, self.font_size);
        style.bold = self.bold;
        style.italic = self.italic;
        style.underline = self.underline;
        style.strikethrough = self.strikethrough || self.double_strikethrough;

        if let Some(color) = Color::from_hex(&self.color) {
            style.color = color.to_rgba_array();
        }

        // Combine highlight and shading (highlight takes precedence)
        let bg_color = self.highlight.as_ref().or(self.shading.as_ref());
        if let Some(bg) = bg_color {
            if let Some(color) = Color::from_hex(bg) {
                style.background = Some(color.to_rgba_array());
            }
        }

        style
    }

    /// Convert to TextRun for layout.
    pub fn to_text_run(&self) -> TextRun {
        TextRun::new(&self.text, self.to_text_style())
    }

    /// Get the effective font size accounting for superscript/subscript.
    pub fn effective_font_size(&self) -> f32 {
        match self.vertical_align.as_str() {
            "superscript" | "subscript" => self.font_size * 0.65,
            _ => self.font_size,
        }
    }

    /// Get the vertical offset for super/subscript rendering.
    pub fn vertical_offset(&self) -> f32 {
        match self.vertical_align.as_str() {
            "superscript" => -self.font_size * 0.35,
            "subscript" => self.font_size * 0.15,
            _ => self.position,
        }
    }

    /// Apply text transformations (caps, small caps).
    pub fn transformed_text(&self) -> String {
        if self.all_caps {
            self.text.to_uppercase()
        } else if self.small_caps {
            // For small caps, we'd ideally use CSS font-variant, but for rendering
            // we'll just use uppercase and reduce size
            self.text.to_uppercase()
        } else {
            self.text.clone()
        }
    }
}

/// Border definition for paragraph and table borders.
#[derive(Clone, Debug, Serialize, Deserialize, Default)]
pub struct ParagraphBorder {
    /// Border style (none, single, double, dotted, dashed, etc.)
    pub style: String,
    /// Border width in points
    pub width: f32,
    /// Border color (hex)
    pub color: String,
    /// Space between border and content in points
    pub space: f32,
}

/// A styled paragraph extracted from DOCX with comprehensive formatting.
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct StyledParagraph {
    /// Text runs in this paragraph
    pub runs: Vec<StyledRun>,
    /// Text alignment: "left", "center", "right", "both" (justify), "distribute"
    pub alignment: String,
    /// Line spacing multiplier (1.0 = single, 1.5 = one-and-a-half, 2.0 = double)
    pub line_spacing: f32,
    /// Line spacing rule: "auto", "exact", "atLeast"
    #[serde(default)]
    pub line_spacing_rule: String,
    /// Space before paragraph in points
    pub space_before: f32,
    /// Space after paragraph in points
    pub space_after: f32,
    /// First line indent in points (positive = indent, negative with hanging = hanging indent)
    pub first_line_indent: f32,
    /// Left indent in points
    pub left_indent: f32,
    /// Right indent in points
    pub right_indent: f32,
    /// Hanging indent in points (used for lists)
    #[serde(default)]
    pub hanging_indent: f32,
    /// Is this a heading?
    pub is_heading: bool,
    /// Heading level (1-9 for Heading1-9)
    pub heading_level: Option<u8>,
    /// Outline level (for document outline)
    #[serde(default)]
    pub outline_level: Option<u8>,
    /// Style ID (e.g., "Heading1", "Normal")
    #[serde(default)]
    pub style_id: Option<String>,
    /// Paragraph background/shading color (hex)
    #[serde(default)]
    pub shading: Option<String>,
    /// Top border
    #[serde(default)]
    pub border_top: Option<ParagraphBorder>,
    /// Bottom border
    #[serde(default)]
    pub border_bottom: Option<ParagraphBorder>,
    /// Left border
    #[serde(default)]
    pub border_left: Option<ParagraphBorder>,
    /// Right border
    #[serde(default)]
    pub border_right: Option<ParagraphBorder>,
    /// Border between paragraphs (for consecutive bordered paragraphs)
    #[serde(default)]
    pub border_between: Option<ParagraphBorder>,
    /// Numbering/list ID (-1 if not a list item)
    #[serde(default)]
    pub numbering_id: Option<i32>,
    /// Numbering level (0-8)
    #[serde(default)]
    pub numbering_level: Option<i32>,
    /// Pre-formatted list prefix (e.g., "1.", "•", "a)")
    #[serde(default)]
    pub list_prefix: Option<String>,
    /// Keep with next paragraph
    #[serde(default)]
    pub keep_next: bool,
    /// Keep lines together (no page break within)
    #[serde(default)]
    pub keep_lines: bool,
    /// Page break before this paragraph
    #[serde(default)]
    pub page_break_before: bool,
    /// Widow/orphan control
    #[serde(default)]
    pub widow_control: bool,
    /// Tab stops
    #[serde(default)]
    pub tab_stops: Vec<TabStopDef>,
    /// Is this paragraph part of a footnote?
    #[serde(default)]
    pub is_footnote: bool,
    /// Footnote reference number (if footnote)
    #[serde(default)]
    pub footnote_number: Option<i32>,
    /// Is this paragraph part of an endnote?
    #[serde(default)]
    pub is_endnote: bool,
    /// Endnote reference number (if endnote)
    #[serde(default)]
    pub endnote_number: Option<i32>,
}

/// Tab stop definition.
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct TabStopDef {
    /// Position in points
    pub position: f32,
    /// Alignment: "left", "center", "right", "decimal", "bar"
    pub alignment: String,
    /// Leader character: "none", "dot", "hyphen", "underscore"
    pub leader: String,
}

impl Default for StyledParagraph {
    fn default() -> Self {
        Self {
            runs: Vec::new(),
            alignment: "left".to_string(),
            line_spacing: 1.15,
            line_spacing_rule: "auto".to_string(),
            space_before: 0.0,
            space_after: 8.0,
            first_line_indent: 0.0,
            left_indent: 0.0,
            right_indent: 0.0,
            hanging_indent: 0.0,
            is_heading: false,
            heading_level: None,
            outline_level: None,
            style_id: None,
            shading: None,
            border_top: None,
            border_bottom: None,
            border_left: None,
            border_right: None,
            border_between: None,
            numbering_id: None,
            numbering_level: None,
            list_prefix: None,
            keep_next: false,
            keep_lines: false,
            page_break_before: false,
            widow_control: true,
            tab_stops: Vec::new(),
            is_footnote: false,
            footnote_number: None,
            is_endnote: false,
            endnote_number: None,
        }
    }
}

impl StyledParagraph {
    /// Convert to Paragraph for text layout.
    pub fn to_paragraph(&self) -> Paragraph {
        let runs: Vec<TextRun> = self.runs.iter().map(|r| r.to_text_run()).collect();
        let align = match self.alignment.as_str() {
            "center" => TextAlign::Center,
            "right" => TextAlign::Right,
            "justify" | "both" => TextAlign::Justify,
            "distribute" => TextAlign::Justify, // Approximate
            _ => TextAlign::Left,
        };

        // Calculate effective first line indent (accounting for hanging indent)
        let effective_first_line = if self.hanging_indent > 0.0 {
            -self.hanging_indent
        } else {
            self.first_line_indent
        };

        Paragraph {
            runs,
            align,
            line_spacing: self.line_spacing,
            space_before: self.space_before,
            space_after: self.space_after,
            first_line_indent: effective_first_line,
            left_indent: self.left_indent + self.hanging_indent.max(0.0),
            right_indent: self.right_indent,
        }
    }

    /// Get the combined text of all runs.
    pub fn get_text(&self) -> String {
        self.runs.iter().map(|r| r.text.as_str()).collect()
    }

    /// Check if this paragraph has any visible borders.
    pub fn has_borders(&self) -> bool {
        self.border_top
            .as_ref()
            .map(|b| b.style != "none")
            .unwrap_or(false)
            || self
                .border_bottom
                .as_ref()
                .map(|b| b.style != "none")
                .unwrap_or(false)
            || self
                .border_left
                .as_ref()
                .map(|b| b.style != "none")
                .unwrap_or(false)
            || self
                .border_right
                .as_ref()
                .map(|b| b.style != "none")
                .unwrap_or(false)
    }

    /// Get effective line height in points based on the first run's font size.
    pub fn line_height(&self) -> f32 {
        let base_size = self.runs.first().map(|r| r.font_size).unwrap_or(11.0);

        match self.line_spacing_rule.as_str() {
            "exact" => self.line_spacing, // Already in points
            "atLeast" => self.line_spacing.max(base_size * 1.2),
            _ => base_size * self.line_spacing * 1.2, // Auto: multiplier
        }
    }

    /// Check if this is a list item.
    pub fn is_list_item(&self) -> bool {
        self.numbering_id.is_some() && self.numbering_id != Some(-1)
    }
}

/// Table cell border definition.
#[derive(Clone, Debug, Serialize, Deserialize, Default)]
pub struct CellBorder {
    /// Border style
    pub style: String,
    /// Border width in points
    pub width: f32,
    /// Border color (hex)
    pub color: String,
}

/// Table cell from DOCX with comprehensive properties.
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct TableCell {
    /// Paragraphs within the cell
    pub paragraphs: Vec<StyledParagraph>,
    /// Column span (horizontal merge)
    pub col_span: u32,
    /// Row span (vertical merge tracking)
    #[serde(default)]
    pub row_span: u32,
    /// Is this cell a continuation of a vertical merge?
    #[serde(default)]
    pub is_v_merge_continue: bool,
    /// Cell width in points
    pub width: Option<f32>,
    /// Background/shading color (hex)
    pub background: Option<String>,
    /// Top border
    #[serde(default)]
    pub border_top: Option<CellBorder>,
    /// Bottom border
    #[serde(default)]
    pub border_bottom: Option<CellBorder>,
    /// Left border
    #[serde(default)]
    pub border_left: Option<CellBorder>,
    /// Right border
    #[serde(default)]
    pub border_right: Option<CellBorder>,
    /// Vertical alignment: "top", "center", "bottom"
    #[serde(default)]
    pub vertical_align: String,
    /// Text direction: "lrTb" (left-to-right, top-to-bottom), "tbRl", "btLr"
    #[serde(default)]
    pub text_direction: String,
    /// Cell margins (top, right, bottom, left) in points
    #[serde(default)]
    pub margins: (f32, f32, f32, f32),
    /// No text wrap
    #[serde(default)]
    pub no_wrap: bool,
}

impl Default for TableCell {
    fn default() -> Self {
        Self {
            paragraphs: Vec::new(),
            col_span: 1,
            row_span: 1,
            is_v_merge_continue: false,
            width: None,
            background: None,
            border_top: None,
            border_bottom: None,
            border_left: None,
            border_right: None,
            vertical_align: "top".to_string(),
            text_direction: "lrTb".to_string(),
            margins: (0.0, 5.4, 0.0, 5.4), // Default cell margins (approx 0.08")
            no_wrap: false,
        }
    }
}

/// Table row from DOCX with properties.
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct TableRow {
    /// Cells in this row
    pub cells: Vec<TableCell>,
    /// Row height in points (None = auto)
    #[serde(default)]
    pub height: Option<f32>,
    /// Height rule: "auto", "exact", "atLeast"
    #[serde(default)]
    pub height_rule: String,
    /// Is this a header row (repeats at top of each page)?
    #[serde(default)]
    pub is_header: bool,
    /// Can this row split across pages?
    #[serde(default)]
    pub can_split: bool,
}

impl Default for TableRow {
    fn default() -> Self {
        Self {
            cells: Vec::new(),
            height: None,
            height_rule: "auto".to_string(),
            is_header: false,
            can_split: true,
        }
    }
}

/// Table from DOCX with comprehensive properties.
#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct Table {
    /// Rows in the table
    pub rows: Vec<TableRow>,
    /// Column widths in points
    pub column_widths: Vec<f32>,
    /// Table width in points (None = auto)
    #[serde(default)]
    pub width: Option<f32>,
    /// Table alignment: "left", "center", "right"
    #[serde(default)]
    pub alignment: String,
    /// Table indent from left margin in points
    #[serde(default)]
    pub indent: f32,
    /// Default cell margins (top, right, bottom, left) in points
    #[serde(default)]
    pub default_cell_margins: (f32, f32, f32, f32),
    /// Cell spacing in points
    #[serde(default)]
    pub cell_spacing: f32,
    /// Table borders
    #[serde(default)]
    pub border_top: Option<CellBorder>,
    #[serde(default)]
    pub border_bottom: Option<CellBorder>,
    #[serde(default)]
    pub border_left: Option<CellBorder>,
    #[serde(default)]
    pub border_right: Option<CellBorder>,
    #[serde(default)]
    pub border_inside_h: Option<CellBorder>,
    #[serde(default)]
    pub border_inside_v: Option<CellBorder>,
    /// Table style ID
    #[serde(default)]
    pub style_id: Option<String>,
    /// Table look flags (for conditional formatting)
    #[serde(default)]
    pub first_row_style: bool,
    #[serde(default)]
    pub last_row_style: bool,
    #[serde(default)]
    pub first_col_style: bool,
    #[serde(default)]
    pub last_col_style: bool,
    #[serde(default)]
    pub banded_rows: bool,
    #[serde(default)]
    pub banded_cols: bool,
}

impl Default for Table {
    fn default() -> Self {
        Self {
            rows: Vec::new(),
            column_widths: Vec::new(),
            width: None,
            alignment: "left".to_string(),
            indent: 0.0,
            default_cell_margins: (0.0, 5.4, 0.0, 5.4),
            cell_spacing: 0.0,
            border_top: None,
            border_bottom: None,
            border_left: None,
            border_right: None,
            border_inside_h: None,
            border_inside_v: None,
            style_id: None,
            first_row_style: false,
            last_row_style: false,
            first_col_style: false,
            last_col_style: false,
            banded_rows: false,
            banded_cols: false,
        }
    }
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
    /// Header margin in points
    header_margin: f32,
    /// Footer margin in points
    footer_margin: f32,
    /// Media files (images) keyed by relationship ID
    media: std::collections::HashMap<String, (Vec<u8>, String)>,
    /// Style resolver for comprehensive style resolution
    style_resolver: Option<StyleResolver>,
    /// Numbering counters (for tracking list numbers)
    numbering_counters: HashMap<(i32, i32), i32>, // (num_id, level) -> current count
}

impl DocxDocument {
    /// Parse a DOCX document from bytes with comprehensive style resolution.
    pub fn from_bytes(data: &[u8]) -> Result<Self, DocxError> {
        let docx = read_docx(data)
            .map_err(|e| DocxError::ParseError(format!("Failed to parse DOCX: {:?}", e)))?;

        // Create style resolver for comprehensive style handling
        let style_resolver = StyleResolver::new(&docx);

        // Extract page size and margins from section properties
        let (
            page_width,
            page_height,
            margin_top,
            margin_right,
            margin_bottom,
            margin_left,
            header_margin,
            footer_margin,
        ) = Self::extract_page_settings(&docx);

        // Build media map from document relationships and media files
        let mut media = std::collections::HashMap::new();

        // Extract media files from docx
        for (filename, data) in &docx.media {
            let format = filename.rsplit('.').next().unwrap_or("png").to_lowercase();
            media.insert(filename.clone(), (data.clone(), format));
        }

        // Also check the images field which has (id, path, image, png) tuples
        for (id, path, _image, png) in &docx.images {
            let format = path.rsplit('.').next().unwrap_or("png").to_lowercase();
            media.insert(id.clone(), (png.0.clone(), format));
        }

        // Initialize numbering counters
        let mut numbering_counters: HashMap<(i32, i32), i32> = HashMap::new();

        // Extract document elements with full style resolution
        let mut elements = Vec::new();

        for child in docx.document.children.iter() {
            match child {
                DocumentChild::Paragraph(para) => {
                    // Check for images in the paragraph
                    let images = Self::extract_images_from_paragraph(para, &media);
                    for img in images {
                        elements.push(DocumentElement::Image(img));
                    }

                    // Extract paragraph with full style resolution
                    let styled = Self::extract_paragraph_with_styles(
                        para,
                        &style_resolver,
                        &mut numbering_counters,
                    );

                    // Add paragraph if it has content, or if it's a heading, or if it has borders/shading
                    if !styled.runs.is_empty()
                        || styled.is_heading
                        || styled.has_borders()
                        || styled.shading.is_some()
                    {
                        elements.push(DocumentElement::Paragraph(styled));
                    }
                }
                DocumentChild::Table(table) => {
                    let styled = Self::extract_table_with_styles(
                        table,
                        &style_resolver,
                        &mut numbering_counters,
                    );
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
            header_margin,
            footer_margin,
            media,
            style_resolver: Some(style_resolver),
            numbering_counters,
        })
    }

    /// Extract page settings from section properties.
    fn extract_page_settings(docx: &Docx) -> (f32, f32, f32, f32, f32, f32, f32, f32) {
        // Default values (US Letter)
        let mut page_width = 612.0; // 8.5 inches
        let mut page_height = 792.0; // 11 inches
        let mut margin_top = 72.0; // 1 inch
        let mut margin_right = 72.0;
        let mut margin_bottom = 72.0;
        let mut margin_left = 72.0;
        let mut header_margin = 36.0; // 0.5 inch
        let mut footer_margin = 36.0;

        // Try to extract from section properties via JSON serialization
        if let Ok(json) = serde_json::to_string(&docx.document.section_property) {
            if let Ok(props) = serde_json::from_str::<serde_json::Value>(&json) {
                // Page size
                if let Some(page_size) = props.get("pageSize") {
                    if let Some(w) = page_size.get("w").and_then(|v| v.as_i64()) {
                        page_width = w as f32 * TWIPS_TO_POINTS;
                    }
                    if let Some(h) = page_size.get("h").and_then(|v| v.as_i64()) {
                        page_height = h as f32 * TWIPS_TO_POINTS;
                    }
                }

                // Margins
                if let Some(margins) = props.get("pageMargin") {
                    if let Some(top) = margins.get("top").and_then(|v| v.as_i64()) {
                        margin_top = top as f32 * TWIPS_TO_POINTS;
                    }
                    if let Some(right) = margins.get("right").and_then(|v| v.as_i64()) {
                        margin_right = right as f32 * TWIPS_TO_POINTS;
                    }
                    if let Some(bottom) = margins.get("bottom").and_then(|v| v.as_i64()) {
                        margin_bottom = bottom as f32 * TWIPS_TO_POINTS;
                    }
                    if let Some(left) = margins.get("left").and_then(|v| v.as_i64()) {
                        margin_left = left as f32 * TWIPS_TO_POINTS;
                    }
                    if let Some(header) = margins.get("header").and_then(|v| v.as_i64()) {
                        header_margin = header as f32 * TWIPS_TO_POINTS;
                    }
                    if let Some(footer) = margins.get("footer").and_then(|v| v.as_i64()) {
                        footer_margin = footer as f32 * TWIPS_TO_POINTS;
                    }
                }
            }
        }

        (
            page_width,
            page_height,
            margin_top,
            margin_right,
            margin_bottom,
            margin_left,
            header_margin,
            footer_margin,
        )
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

    /// Extract paragraph with full style resolution.
    fn extract_paragraph_with_styles(
        para: &docx_rs::Paragraph,
        style_resolver: &StyleResolver,
        numbering_counters: &mut HashMap<(i32, i32), i32>,
    ) -> StyledParagraph {
        let mut styled = StyledParagraph::default();
        let props = &para.property;

        // Get the paragraph style ID
        let style_id = props.style.as_ref().map(|s| s.val.as_str());
        styled.style_id = style_id.map(|s| s.to_string());

        // First, extract direct paragraph properties as ResolvedParaProps
        let direct_props = Self::extract_direct_para_props(props);

        // Resolve complete paragraph properties with style inheritance
        let resolved_props = style_resolver.resolve_para_props(style_id, &direct_props);

        // Apply resolved properties to styled paragraph
        styled.alignment = match resolved_props
            .alignment
            .clone()
            .unwrap_or(AlignmentType::Left)
        {
            AlignmentType::Left => "left".to_string(),
            AlignmentType::Center => "center".to_string(),
            AlignmentType::Right => "right".to_string(),
            AlignmentType::Both => "both".to_string(),
            AlignmentType::Distribute => "distribute".to_string(),
        };

        styled.left_indent = resolved_props.left_indent.unwrap_or(0.0);
        styled.right_indent = resolved_props.right_indent.unwrap_or(0.0);
        styled.first_line_indent = resolved_props.first_line_indent.unwrap_or(0.0);
        styled.hanging_indent = resolved_props.hanging_indent.unwrap_or(0.0);
        styled.space_before = resolved_props.space_before.unwrap_or(0.0);
        styled.space_after = resolved_props.space_after.unwrap_or(8.0);
        styled.line_spacing = resolved_props.line_spacing.unwrap_or(1.15);
        styled.line_spacing_rule = match resolved_props
            .line_spacing_rule
            .clone()
            .unwrap_or(LineSpacingRule::Auto)
        {
            LineSpacingRule::Auto => "auto".to_string(),
            LineSpacingRule::Exact => "exact".to_string(),
            LineSpacingRule::AtLeast => "atLeast".to_string(),
        };

        styled.shading = resolved_props.shading.clone();
        styled.keep_next = resolved_props.keep_next.unwrap_or(false);
        styled.keep_lines = resolved_props.keep_lines.unwrap_or(false);
        styled.page_break_before = resolved_props.page_break_before.unwrap_or(false);
        styled.widow_control = resolved_props.widow_control.unwrap_or(true);
        styled.outline_level = resolved_props.outline_level;

        // Convert borders
        styled.border_top = resolved_props
            .border_top
            .as_ref()
            .map(|b| Self::convert_border(b));
        styled.border_bottom = resolved_props
            .border_bottom
            .as_ref()
            .map(|b| Self::convert_border(b));
        styled.border_left = resolved_props
            .border_left
            .as_ref()
            .map(|b| Self::convert_border(b));
        styled.border_right = resolved_props
            .border_right
            .as_ref()
            .map(|b| Self::convert_border(b));
        styled.border_left = resolved_props
            .border_left
            .as_ref()
            .map(|b| Self::convert_border(b));
        styled.border_right = resolved_props
            .border_right
            .as_ref()
            .map(|b| Self::convert_border(b));
        styled.border_between = resolved_props
            .border_between
            .as_ref()
            .map(|b| Self::convert_border(b));

        // Check for heading style
        if let Some(ref sid) = styled.style_id {
            if sid.starts_with("Heading") || sid.starts_with("heading") {
                styled.is_heading = true;
                styled.heading_level = sid
                    .chars()
                    .last()
                    .and_then(|c| c.to_digit(10).map(|d| d as u8));
            }
            // Also check outline level
            if let Some(level) = styled.outline_level {
                styled.is_heading = true;
                styled.heading_level = Some(level + 1);
            }
        }

        // Handle numbering/lists
        styled.numbering_id = resolved_props.numbering_id;
        styled.numbering_level = resolved_props.numbering_level;

        if let (Some(num_id), Some(level)) = (styled.numbering_id, styled.numbering_level) {
            if let Some(numbering_def) = style_resolver.get_numbering(num_id) {
                if let Some(level_def) = numbering_def.levels.get(level as usize) {
                    // Update counter
                    let counter_key = (num_id, level);
                    let counter = numbering_counters
                        .entry(counter_key)
                        .or_insert(level_def.start - 1);
                    *counter += 1;

                    // Format the list prefix
                    let formatted_num = style_resolver.format_number(*counter, &level_def.num_fmt);
                    let prefix = level_def
                        .lvl_text
                        .replace(&format!("%{}", level + 1), &formatted_num);
                    styled.list_prefix = Some(prefix);

                    // Apply numbering indent
                    if level_def.indent_left > 0.0 {
                        styled.left_indent = level_def.indent_left;
                    }
                    if level_def.hanging > 0.0 {
                        styled.hanging_indent = level_def.hanging;
                    }
                }
            }
        }

        // Extract runs with full style resolution
        for child in &para.children {
            match child {
                ParagraphChild::Run(run) => {
                    let styled_run = Self::extract_run_with_styles(
                        run,
                        style_id,
                        style_resolver,
                        &resolved_props,
                    );
                    if !styled_run.text.is_empty() {
                        styled.runs.push(styled_run);
                    }
                }
                ParagraphChild::Hyperlink(hyperlink) => {
                    // Extract runs from hyperlink
                    for hlink_child in &hyperlink.children {
                        if let ParagraphChild::Run(run) = hlink_child {
                            let mut styled_run = Self::extract_run_with_styles(
                                run,
                                style_id,
                                style_resolver,
                                &resolved_props,
                            );
                            styled_run.is_hyperlink = true;
                            // Try to get the hyperlink URL
                            if let Ok(json) = serde_json::to_string(&hyperlink) {
                                if let Ok(parsed) = serde_json::from_str::<serde_json::Value>(&json)
                                {
                                    if let Some(rid) = parsed.get("rid").and_then(|r| r.as_str()) {
                                        styled_run.hyperlink_url = Some(rid.to_string());
                                    }
                                }
                            }
                            if !styled_run.text.is_empty() {
                                styled.runs.push(styled_run);
                            }
                        }
                    }
                }
                ParagraphChild::BookmarkStart(bookmark) => {
                    // Note: Bookmark starts don't have text, but we track the name
                    if let Ok(json) = serde_json::to_string(bookmark) {
                        if let Ok(parsed) = serde_json::from_str::<serde_json::Value>(&json) {
                            if let Some(name) = parsed.get("name").and_then(|n| n.as_str()) {
                                // Mark the next run with this bookmark
                                // For now, just log the bookmark name
                                let _ = name;
                            }
                        }
                    }
                }
                _ => {}
            }
        }

        styled
    }

    /// Extract direct paragraph properties (without style resolution).
    fn extract_direct_para_props(props: &docx_rs::ParagraphProperty) -> ResolvedParaProps {
        let mut resolved = ResolvedParaProps::default();

        // Serialize the entire property to JSON for extraction
        if let Ok(json) = serde_json::to_string(props) {
            if let Ok(parsed) = serde_json::from_str::<serde_json::Value>(&json) {
                // Alignment
                if let Some(align) = parsed.get("alignment").and_then(|a| a.as_str()) {
                    resolved.alignment = Some(match align.to_lowercase().as_str() {
                        "left" | "start" => AlignmentType::Left,
                        "center" => AlignmentType::Center,
                        "right" | "end" => AlignmentType::Right,
                        "both" | "justify" => AlignmentType::Both,
                        "distribute" => AlignmentType::Distribute,
                        _ => AlignmentType::Left,
                    });
                }

                // Indentation
                if let Some(indent) = parsed.get("indent") {
                    if let Some(left) = indent.get("start").and_then(|s| s.as_i64()) {
                        resolved.left_indent = Some(left as f32 * TWIPS_TO_POINTS);
                    } else if let Some(left) = indent.get("left").and_then(|s| s.as_i64()) {
                        resolved.left_indent = Some(left as f32 * TWIPS_TO_POINTS);
                    }
                    if let Some(right) = indent.get("end").and_then(|e| e.as_i64()) {
                        resolved.right_indent = Some(right as f32 * TWIPS_TO_POINTS);
                    } else if let Some(right) = indent.get("right").and_then(|e| e.as_i64()) {
                        resolved.right_indent = Some(right as f32 * TWIPS_TO_POINTS);
                    }
                    if let Some(first) = indent.get("firstLine").and_then(|f| f.as_i64()) {
                        resolved.first_line_indent = Some(first as f32 * TWIPS_TO_POINTS);
                    }
                    if let Some(hanging) = indent.get("hanging").and_then(|h| h.as_i64()) {
                        resolved.hanging_indent = Some(hanging as f32 * TWIPS_TO_POINTS);
                    }
                }

                // Spacing
                if let Some(spacing) = parsed.get("lineSpacing") {
                    if let Some(before) = spacing.get("before").and_then(|b| b.as_i64()) {
                        resolved.space_before = Some(before as f32 * TWIPS_TO_POINTS);
                    }
                    if let Some(after) = spacing.get("after").and_then(|a| a.as_i64()) {
                        resolved.space_after = Some(after as f32 * TWIPS_TO_POINTS);
                    }
                    if let Some(line) = spacing.get("line").and_then(|l| l.as_i64()) {
                        let rule = spacing
                            .get("lineRule")
                            .and_then(|r| r.as_str())
                            .unwrap_or("auto");
                        resolved.line_spacing_rule = Some(match rule {
                            "exact" => LineSpacingRule::Exact,
                            "atLeast" => LineSpacingRule::AtLeast,
                            _ => LineSpacingRule::Auto,
                        });
                        match resolved.line_spacing_rule {
                            Some(LineSpacingRule::Auto) => {
                                resolved.line_spacing = Some(line as f32 / 240.0);
                            }
                            _ => {
                                resolved.line_spacing = Some(line as f32 * TWIPS_TO_POINTS);
                            }
                        }
                    }
                }

                // Shading
                if let Some(shading) = parsed.get("shading") {
                    if let Some(fill) = shading.get("fill").and_then(|f| f.as_str()) {
                        if fill != "auto" && !fill.is_empty() {
                            resolved.shading = Some(format!("#{}", fill));
                        }
                    }
                }

                // Borders
                if let Some(borders) = parsed.get("borders") {
                    resolved.border_top = Self::extract_border_from_json(borders.get("top"));
                    resolved.border_bottom = Self::extract_border_from_json(borders.get("bottom"));
                    resolved.border_left = Self::extract_border_from_json(borders.get("left"));
                    resolved.border_right = Self::extract_border_from_json(borders.get("right"));
                    resolved.border_between =
                        Self::extract_border_from_json(borders.get("between"));
                }

                // Numbering
                if let Some(num_pr) = parsed.get("numberingProperty") {
                    if let Some(num_id) = num_pr.get("id").and_then(|i| i.as_i64()) {
                        resolved.numbering_id = Some(num_id as i32);
                    }
                    if let Some(ilvl) = num_pr.get("level").and_then(|l| l.as_i64()) {
                        resolved.numbering_level = Some(ilvl as i32);
                    }
                }

                // Other properties
                if parsed.get("keepNext").is_some() {
                    resolved.keep_next = Some(true);
                }
                if parsed.get("keepLines").is_some() {
                    resolved.keep_lines = Some(true);
                }
                if parsed.get("pageBreakBefore").is_some() {
                    resolved.page_break_before = Some(true);
                }
                if parsed.get("widowControl").is_some() {
                    resolved.widow_control = Some(true);
                }
                if let Some(outline_lvl) = parsed
                    .get("outlineLvl")
                    .and_then(|o| o.get("val"))
                    .and_then(|v| v.as_i64())
                {
                    resolved.outline_level = Some(outline_lvl as u8);
                }
            }
        }

        resolved
    }

    /// Extract border definition from JSON.
    fn extract_border_from_json(border: Option<&serde_json::Value>) -> Option<BorderDef> {
        let border = border?;
        let style_str = border.get("val").and_then(|v| v.as_str()).unwrap_or("none");

        if style_str == "nil" || style_str == "none" {
            return None;
        }

        let style = match style_str {
            "single" => BorderStyleType::Single,
            "thick" => BorderStyleType::Thick,
            "double" => BorderStyleType::Double,
            "dotted" => BorderStyleType::Dotted,
            "dashed" => BorderStyleType::Dashed,
            "dotDash" => BorderStyleType::DotDash,
            "dotDotDash" => BorderStyleType::DotDotDash,
            "triple" => BorderStyleType::Triple,
            "wave" => BorderStyleType::Wave,
            _ => BorderStyleType::Single,
        };

        let width = border
            .get("sz")
            .and_then(|s| s.as_i64())
            .map(|s| s as f32 * EIGHTHS_TO_POINTS)
            .unwrap_or(1.0);

        let color = border
            .get("color")
            .and_then(|c| c.as_str())
            .map(|c| {
                if c == "auto" {
                    "#000000".to_string()
                } else {
                    format!("#{}", c)
                }
            })
            .unwrap_or_else(|| "#000000".to_string());

        let space = border
            .get("space")
            .and_then(|s| s.as_i64())
            .map(|s| s as f32)
            .unwrap_or(0.0);

        Some(BorderDef {
            style,
            width,
            color,
            space,
        })
    }

    /// Convert BorderDef to ParagraphBorder for storage.
    fn convert_border(def: &BorderDef) -> ParagraphBorder {
        ParagraphBorder {
            style: match def.style {
                BorderStyleType::None => "none",
                BorderStyleType::Single => "single",
                BorderStyleType::Thick => "thick",
                BorderStyleType::Double => "double",
                BorderStyleType::Dotted => "dotted",
                BorderStyleType::Dashed => "dashed",
                BorderStyleType::DotDash => "dotDash",
                BorderStyleType::DotDotDash => "dotDotDash",
                BorderStyleType::Triple => "triple",
                BorderStyleType::Wave => "wave",
                BorderStyleType::DoubleWave => "doubleWave",
                _ => "single",
            }
            .to_string(),
            width: def.width,
            color: def.color.clone(),
            space: def.space,
        }
    }

    /// Extract run with full style resolution.
    fn extract_run_with_styles(
        run: &docx_rs::Run,
        para_style_id: Option<&str>,
        style_resolver: &StyleResolver,
        para_props: &ResolvedParaProps,
    ) -> StyledRun {
        let mut styled = StyledRun::default();

        // Extract text content
        for child in &run.children {
            match child {
                RunChild::Text(t) => {
                    styled.text.push_str(&t.text);
                }
                RunChild::Tab(_) => {
                    styled.text.push('\t');
                }
                RunChild::Break(br) => {
                    // Check break type via JSON
                    if let Ok(json) = serde_json::to_string(br) {
                        if json.contains("page") {
                            styled.text.push_str("\x0C"); // Form feed for page break
                        } else {
                            styled.text.push('\n');
                        }
                    } else {
                        styled.text.push('\n');
                    }
                }
                RunChild::Sym(sym) => {
                    // Symbol characters
                    if let Ok(json) = serde_json::to_string(sym) {
                        if let Ok(parsed) = serde_json::from_str::<serde_json::Value>(&json) {
                            if let Some(char_code) = parsed.get("char").and_then(|c| c.as_str()) {
                                if let Ok(code) =
                                    u32::from_str_radix(char_code.trim_start_matches("F0"), 16)
                                {
                                    if let Some(ch) = char::from_u32(code) {
                                        styled.text.push(ch);
                                    }
                                }
                            }
                        }
                    }
                }
                _ => {}
            }
        }

        // Extract direct run properties
        let direct_props = Self::extract_direct_run_props(&run.run_property, style_resolver);

        // Get character style ID if any
        let char_style_id = if let Ok(json) = serde_json::to_string(&run.run_property) {
            serde_json::from_str::<serde_json::Value>(&json)
                .ok()
                .and_then(|p| {
                    p.get("style")
                        .and_then(|s| s.get("val"))
                        .and_then(|v| v.as_str())
                        .map(|s| s.to_string())
                })
        } else {
            None
        };

        // Resolve run properties with full inheritance
        let resolved = style_resolver.resolve_run_props(
            para_style_id,
            char_style_id.as_deref(),
            &direct_props,
        );

        // Also apply paragraph's default run properties for any remaining defaults
        let final_resolved =
            StyleResolver::merge_run_props(&para_props.default_run_props, &resolved);

        // Apply resolved properties to styled run
        styled.font_family = final_resolved
            .font_family
            .unwrap_or_else(|| "Calibri".to_string());
        styled.font_size = final_resolved.font_size.unwrap_or(11.0);
        styled.bold = final_resolved.bold.unwrap_or(false);
        styled.italic = final_resolved.italic.unwrap_or(false);
        styled.underline = final_resolved
            .underline
            .as_ref()
            .map(|u| *u != UnderlineStyle::None)
            .unwrap_or(false);
        styled.underline_style = final_resolved
            .underline
            .as_ref()
            .map(|u| match u {
                UnderlineStyle::None => "none",
                UnderlineStyle::Single => "single",
                UnderlineStyle::Words => "words",
                UnderlineStyle::Double => "double",
                UnderlineStyle::Thick => "thick",
                UnderlineStyle::Dotted => "dotted",
                UnderlineStyle::Dash => "dash",
                UnderlineStyle::Wave => "wave",
                _ => "single",
            })
            .unwrap_or("none")
            .to_string();
        styled.strikethrough = final_resolved.strike.unwrap_or(false);
        styled.double_strikethrough = final_resolved.double_strike.unwrap_or(false);
        styled.color = final_resolved
            .color
            .unwrap_or_else(|| "#000000".to_string());
        styled.highlight = final_resolved.highlight;
        styled.shading = final_resolved.shading;
        styled.vertical_align = match final_resolved
            .vert_align
            .unwrap_or(VerticalAlignType::Baseline)
        {
            VerticalAlignType::Baseline => "baseline",
            VerticalAlignType::Superscript => "superscript",
            VerticalAlignType::Subscript => "subscript",
        }
        .to_string();
        styled.all_caps = final_resolved.caps.unwrap_or(false);
        styled.small_caps = final_resolved.small_caps.unwrap_or(false);
        styled.character_spacing = final_resolved.character_spacing.unwrap_or(0.0);
        styled.position = final_resolved.position.unwrap_or(0.0);
        styled.hidden = final_resolved.vanish.unwrap_or(false);
        styled.emboss = final_resolved.emboss.unwrap_or(false);
        styled.imprint = final_resolved.imprint.unwrap_or(false);
        styled.outline = final_resolved.outline.unwrap_or(false);
        styled.shadow = final_resolved.shadow.unwrap_or(false);

        styled
    }

    /// Extract direct run properties from docx-rs RunProperty.
    fn extract_direct_run_props(
        props: &docx_rs::RunProperty,
        style_resolver: &StyleResolver,
    ) -> ResolvedRunProps {
        let mut resolved = ResolvedRunProps::default();

        // Serialize to JSON for comprehensive extraction
        if let Ok(json) = serde_json::to_string(props) {
            if let Ok(parsed) = serde_json::from_str::<serde_json::Value>(&json) {
                // Font size
                if let Some(sz) = parsed.get("sz").and_then(|s| s.as_f64()) {
                    resolved.font_size = Some(sz as f32 * HALF_POINTS_TO_POINTS);
                }

                // Bold
                if parsed.get("bold").is_some() {
                    resolved.bold = Some(true);
                }

                // Italic
                if parsed.get("italic").is_some() {
                    resolved.italic = Some(true);
                }

                // Underline
                if let Some(underline) = parsed.get("underline") {
                    let val = underline
                        .get("val")
                        .and_then(|v| v.as_str())
                        .unwrap_or("single");
                    resolved.underline = Some(match val {
                        "none" => UnderlineStyle::None,
                        "single" => UnderlineStyle::Single,
                        "words" => UnderlineStyle::Words,
                        "double" => UnderlineStyle::Double,
                        "thick" => UnderlineStyle::Thick,
                        "dotted" => UnderlineStyle::Dotted,
                        "dash" => UnderlineStyle::Dash,
                        "wave" => UnderlineStyle::Wave,
                        _ => UnderlineStyle::Single,
                    });
                }

                // Strike-through
                if parsed.get("strike").is_some() {
                    resolved.strike = Some(true);
                }

                // Double strike-through
                if parsed.get("dstrike").is_some() {
                    resolved.double_strike = Some(true);
                }

                // Color - docx-rs serializes Color as just a string "RRGGBB"
                if let Some(color) = parsed.get("color") {
                    if let Some(val) = color.as_str() {
                        // Direct string value
                        if val != "auto" && !val.is_empty() {
                            resolved.color = Some(format!("#{}", val));
                        }
                    } else if let Some(val) = color.get("val").and_then(|v| v.as_str()) {
                        // Fallback: object with val field
                        if val != "auto" && !val.is_empty() {
                            resolved.color = Some(format!("#{}", val));
                        }
                    }
                    // Note: docx-rs doesn't expose theme colors in its Color struct
                    // Theme colors would need to be handled separately if needed
                }

                // Highlight
                if let Some(highlight) = parsed.get("highlight").and_then(|h| h.as_str()) {
                    resolved.highlight = Some(Self::highlight_color_to_hex(highlight));
                }

                // Shading
                if let Some(shading) = parsed.get("shading") {
                    if let Some(fill) = shading.get("fill").and_then(|f| f.as_str()) {
                        if fill != "auto" && !fill.is_empty() {
                            resolved.shading = Some(format!("#{}", fill));
                        }
                    }
                }

                // Vertical alignment
                if let Some(vert_align) = parsed
                    .get("vertAlign")
                    .and_then(|v| v.get("val"))
                    .and_then(|v| v.as_str())
                {
                    resolved.vert_align = Some(match vert_align {
                        "superscript" => VerticalAlignType::Superscript,
                        "subscript" => VerticalAlignType::Subscript,
                        _ => VerticalAlignType::Baseline,
                    });
                }

                // Caps
                if parsed.get("caps").is_some() {
                    resolved.caps = Some(true);
                }

                // Small caps
                if parsed.get("smallCaps").is_some() {
                    resolved.small_caps = Some(true);
                }

                // Character spacing
                if let Some(spacing) = parsed
                    .get("spacing")
                    .and_then(|s| s.get("val"))
                    .and_then(|v| v.as_i64())
                {
                    resolved.character_spacing = Some(spacing as f32 * TWIPS_TO_POINTS);
                }

                // Fonts
                if let Some(fonts) = parsed.get("fonts") {
                    if let Some(ascii) = fonts.get("ascii").and_then(|a| a.as_str()) {
                        resolved.font_family = Some(ascii.to_string());
                    } else if let Some(h_ansi) = fonts.get("hAnsi").and_then(|h| h.as_str()) {
                        resolved.font_family = Some(h_ansi.to_string());
                    } else if let Some(ascii_theme) =
                        fonts.get("asciiTheme").and_then(|a| a.as_str())
                    {
                        resolved.font_family = Some(style_resolver.resolve_theme_font(ascii_theme));
                    }
                }

                // Effects
                if parsed.get("vanish").is_some() {
                    resolved.vanish = Some(true);
                }
                if parsed.get("shadow").is_some() {
                    resolved.shadow = Some(true);
                }
                if parsed.get("outline").is_some() {
                    resolved.outline = Some(true);
                }
                if parsed.get("emboss").is_some() {
                    resolved.emboss = Some(true);
                }
                if parsed.get("imprint").is_some() {
                    resolved.imprint = Some(true);
                }
            }
        }

        resolved
    }

    /// Apply tint or shade to a color.
    fn apply_tint_shade(hex_color: &str, factor: f32, is_tint: bool) -> String {
        let hex = hex_color.trim_start_matches('#');
        if hex.len() < 6 {
            return format!("#{}", hex_color);
        }

        let r = u8::from_str_radix(&hex[0..2], 16).unwrap_or(0) as f32 / 255.0;
        let g = u8::from_str_radix(&hex[2..4], 16).unwrap_or(0) as f32 / 255.0;
        let b = u8::from_str_radix(&hex[4..6], 16).unwrap_or(0) as f32 / 255.0;

        let (r_new, g_new, b_new) = if is_tint {
            // Tint: lighten towards white
            (
                r + (1.0 - r) * factor,
                g + (1.0 - g) * factor,
                b + (1.0 - b) * factor,
            )
        } else {
            // Shade: darken towards black
            (r * factor, g * factor, b * factor)
        };

        format!(
            "#{:02X}{:02X}{:02X}",
            (r_new * 255.0).clamp(0.0, 255.0) as u8,
            (g_new * 255.0).clamp(0.0, 255.0) as u8,
            (b_new * 255.0).clamp(0.0, 255.0) as u8
        )
    }

    /// Convert Word highlight color names to hex colors
    fn highlight_color_to_hex(color_name: &str) -> String {
        match color_name.to_lowercase().as_str() {
            "yellow" => "#FFFF00".to_string(),
            "green" => "#00FF00".to_string(),
            "cyan" => "#00FFFF".to_string(),
            "magenta" => "#FF00FF".to_string(),
            "blue" => "#0000FF".to_string(),
            "red" => "#FF0000".to_string(),
            "darkblue" | "darkBlue" => "#000080".to_string(),
            "darkcyan" | "darkCyan" => "#008080".to_string(),
            "darkgreen" | "darkGreen" => "#008000".to_string(),
            "darkmagenta" | "darkMagenta" => "#800080".to_string(),
            "darkred" | "darkRed" => "#800000".to_string(),
            "darkyellow" | "darkYellow" => "#808000".to_string(),
            "darkgray" | "darkGray" => "#808080".to_string(),
            "lightgray" | "lightGray" => "#C0C0C0".to_string(),
            "black" => "#000000".to_string(),
            "white" => "#FFFFFF".to_string(),
            _ => "#FFFF00".to_string(), // Default to yellow
        }
    }

    /// Extract table with full style resolution.
    fn extract_table_with_styles(
        table: &docx_rs::Table,
        style_resolver: &StyleResolver,
        numbering_counters: &mut HashMap<(i32, i32), i32>,
    ) -> Table {
        let mut result = Table::default();

        // Get column widths from grid
        for width in &table.grid {
            result.column_widths.push(*width as f32 * TWIPS_TO_POINTS);
        }

        // Extract table properties via JSON
        if let Ok(json) = serde_json::to_string(&table.property) {
            if let Ok(parsed) = serde_json::from_str::<serde_json::Value>(&json) {
                // Table width
                if let Some(width) = parsed
                    .get("width")
                    .and_then(|w| w.get("width"))
                    .and_then(|v| v.as_i64())
                {
                    result.width = Some(width as f32 * TWIPS_TO_POINTS);
                }

                // Table alignment
                if let Some(jc) = parsed.get("justification").and_then(|j| j.as_str()) {
                    result.alignment = jc.to_string();
                }

                // Table indent
                if let Some(indent) = parsed
                    .get("indent")
                    .and_then(|i| i.get("width"))
                    .and_then(|v| v.as_i64())
                {
                    result.indent = indent as f32 * TWIPS_TO_POINTS;
                }

                // Cell margins
                if let Some(margins) = parsed.get("cellMargins") {
                    let top = margins
                        .get("top")
                        .and_then(|t| t.get("width"))
                        .and_then(|v| v.as_i64())
                        .unwrap_or(0) as f32
                        * TWIPS_TO_POINTS;
                    let right = margins
                        .get("right")
                        .and_then(|r| r.get("width"))
                        .and_then(|v| v.as_i64())
                        .unwrap_or(108) as f32
                        * TWIPS_TO_POINTS;
                    let bottom = margins
                        .get("bottom")
                        .and_then(|b| b.get("width"))
                        .and_then(|v| v.as_i64())
                        .unwrap_or(0) as f32
                        * TWIPS_TO_POINTS;
                    let left = margins
                        .get("left")
                        .and_then(|l| l.get("width"))
                        .and_then(|v| v.as_i64())
                        .unwrap_or(108) as f32
                        * TWIPS_TO_POINTS;
                    result.default_cell_margins = (top, right, bottom, left);
                }

                // Table borders
                if let Some(borders) = parsed.get("borders") {
                    result.border_top = Self::extract_cell_border_from_json(borders.get("top"));
                    result.border_bottom =
                        Self::extract_cell_border_from_json(borders.get("bottom"));
                    result.border_left = Self::extract_cell_border_from_json(borders.get("left"));
                    result.border_right = Self::extract_cell_border_from_json(borders.get("right"));
                    result.border_inside_h =
                        Self::extract_cell_border_from_json(borders.get("insideH"));
                    result.border_inside_v =
                        Self::extract_cell_border_from_json(borders.get("insideV"));
                }

                // Table style
                if let Some(style) = parsed
                    .get("style")
                    .and_then(|s| s.get("val"))
                    .and_then(|v| v.as_str())
                {
                    result.style_id = Some(style.to_string());
                }

                // Table look (conditional formatting flags)
                if let Some(look) = parsed.get("look") {
                    result.first_row_style = look
                        .get("firstRow")
                        .and_then(|v| v.as_bool())
                        .unwrap_or(false);
                    result.last_row_style = look
                        .get("lastRow")
                        .and_then(|v| v.as_bool())
                        .unwrap_or(false);
                    result.first_col_style = look
                        .get("firstColumn")
                        .and_then(|v| v.as_bool())
                        .unwrap_or(false);
                    result.last_col_style = look
                        .get("lastColumn")
                        .and_then(|v| v.as_bool())
                        .unwrap_or(false);
                    result.banded_rows = look
                        .get("bandedRows")
                        .and_then(|v| v.as_bool())
                        .unwrap_or(false);
                    result.banded_cols = look
                        .get("bandedColumns")
                        .and_then(|v| v.as_bool())
                        .unwrap_or(false);
                }
            }
        }

        // Extract rows
        for (row_idx, table_child) in table.rows.iter().enumerate() {
            let TableChild::TableRow(row) = table_child;
            let mut table_row = TableRow::default();

            // Extract row properties via JSON
            if let Ok(json) = serde_json::to_string(&row.property) {
                if let Ok(parsed) = serde_json::from_str::<serde_json::Value>(&json) {
                    // Row height
                    if let Some(height) = parsed.get("height").and_then(|h| h.as_i64()) {
                        table_row.height = Some(height as f32 * TWIPS_TO_POINTS);
                    }
                    if let Some(rule) = parsed.get("heightRule").and_then(|r| r.as_str()) {
                        table_row.height_rule = rule.to_string();
                    }

                    // Header row
                    if parsed.get("tblHeader").is_some() {
                        table_row.is_header = true;
                    }

                    // Can split
                    if parsed.get("cantSplit").is_some() {
                        table_row.can_split = false;
                    }
                }
            }

            // Is this the first row? Mark as header if table has firstRowStyle
            if row_idx == 0 && result.first_row_style {
                table_row.is_header = true;
            }

            // Extract cells
            for (col_idx, row_child) in row.cells.iter().enumerate() {
                let TableRowChild::TableCell(cell) = row_child;
                let mut table_cell = TableCell::default();

                // Extract cell properties via JSON
                if let Ok(json) = serde_json::to_string(&cell.property) {
                    if let Ok(parsed) = serde_json::from_str::<serde_json::Value>(&json) {
                        // Cell width
                        if let Some(width) = parsed
                            .get("width")
                            .and_then(|w| w.get("width"))
                            .and_then(|v| v.as_i64())
                        {
                            table_cell.width = Some(width as f32 * TWIPS_TO_POINTS);
                        }

                        // Grid span (horizontal merge)
                        if let Some(span) = parsed.get("gridSpan").and_then(|s| s.as_i64()) {
                            table_cell.col_span = span as u32;
                        }

                        // Vertical merge
                        if let Some(v_merge) = parsed.get("vMerge") {
                            if let Some(val) = v_merge.get("val").and_then(|v| v.as_str()) {
                                match val {
                                    "restart" => {
                                        table_cell.row_span = 1; // Will be calculated later
                                    }
                                    "continue" => {
                                        table_cell.is_v_merge_continue = true;
                                    }
                                    _ => {}
                                }
                            } else {
                                // vMerge present without value means continue
                                table_cell.is_v_merge_continue = true;
                            }
                        }

                        // Shading
                        if let Some(shading) = parsed.get("shading") {
                            if let Some(fill) = shading.get("fill").and_then(|f| f.as_str()) {
                                if fill != "auto" && !fill.is_empty() {
                                    table_cell.background = Some(format!("#{}", fill));
                                }
                            }
                            // Also check for theme color
                            if table_cell.background.is_none() {
                                if let Some(theme_fill) =
                                    shading.get("themeFill").and_then(|t| t.as_str())
                                {
                                    let color = style_resolver.resolve_theme_color(theme_fill);
                                    table_cell.background = Some(format!("#{}", color));
                                }
                            }
                        }

                        // Cell borders
                        if let Some(borders) = parsed.get("borders") {
                            table_cell.border_top =
                                Self::extract_cell_border_from_json(borders.get("top"));
                            table_cell.border_bottom =
                                Self::extract_cell_border_from_json(borders.get("bottom"));
                            table_cell.border_left =
                                Self::extract_cell_border_from_json(borders.get("left"));
                            table_cell.border_right =
                                Self::extract_cell_border_from_json(borders.get("right"));
                        }

                        // Vertical alignment
                        if let Some(v_align) = parsed
                            .get("vAlign")
                            .and_then(|v| v.get("val"))
                            .and_then(|v| v.as_str())
                        {
                            table_cell.vertical_align = v_align.to_string();
                        }

                        // Text direction
                        if let Some(text_dir) = parsed
                            .get("textDirection")
                            .and_then(|t| t.get("val"))
                            .and_then(|v| v.as_str())
                        {
                            table_cell.text_direction = text_dir.to_string();
                        }

                        // No wrap
                        if parsed.get("noWrap").is_some() {
                            table_cell.no_wrap = true;
                        }

                        // Cell margins
                        if let Some(margins) = parsed.get("margins") {
                            let top = margins
                                .get("top")
                                .and_then(|t| t.get("width"))
                                .and_then(|v| v.as_i64())
                                .unwrap_or(0) as f32
                                * TWIPS_TO_POINTS;
                            let right = margins
                                .get("right")
                                .and_then(|r| r.get("width"))
                                .and_then(|v| v.as_i64())
                                .unwrap_or(0) as f32
                                * TWIPS_TO_POINTS;
                            let bottom = margins
                                .get("bottom")
                                .and_then(|b| b.get("width"))
                                .and_then(|v| v.as_i64())
                                .unwrap_or(0) as f32
                                * TWIPS_TO_POINTS;
                            let left = margins
                                .get("left")
                                .and_then(|l| l.get("width"))
                                .and_then(|v| v.as_i64())
                                .unwrap_or(0) as f32
                                * TWIPS_TO_POINTS;
                            table_cell.margins = (top, right, bottom, left);
                        } else {
                            table_cell.margins = result.default_cell_margins;
                        }
                    }
                }

                // Apply table-level borders if cell doesn't have its own
                if table_cell.border_top.is_none() && row_idx == 0 {
                    table_cell.border_top = result.border_top.clone();
                }
                if table_cell.border_bottom.is_none() {
                    table_cell.border_bottom = result
                        .border_inside_h
                        .clone()
                        .or_else(|| result.border_bottom.clone());
                }
                if table_cell.border_left.is_none() && col_idx == 0 {
                    table_cell.border_left = result.border_left.clone();
                }
                if table_cell.border_right.is_none() {
                    table_cell.border_right = result
                        .border_inside_v
                        .clone()
                        .or_else(|| result.border_right.clone());
                }

                // Extract cell content (paragraphs)
                for content in &cell.children {
                    if let TableCellContent::Paragraph(para) = content {
                        table_cell
                            .paragraphs
                            .push(Self::extract_paragraph_with_styles(
                                para,
                                style_resolver,
                                numbering_counters,
                            ));
                    }
                }

                table_row.cells.push(table_cell);
            }

            result.rows.push(table_row);
        }

        result
    }

    /// Extract cell border from JSON.
    fn extract_cell_border_from_json(border: Option<&serde_json::Value>) -> Option<CellBorder> {
        let border = border?;
        let style = border.get("val").and_then(|v| v.as_str()).unwrap_or("none");

        if style == "nil" || style == "none" {
            return None;
        }

        let width = border
            .get("sz")
            .and_then(|s| s.as_i64())
            .map(|s| s as f32 * EIGHTHS_TO_POINTS)
            .unwrap_or(1.0);

        let color = border
            .get("color")
            .and_then(|c| c.as_str())
            .map(|c| {
                if c == "auto" {
                    "#000000".to_string()
                } else {
                    format!("#{}", c)
                }
            })
            .unwrap_or_else(|| "#000000".to_string());

        Some(CellBorder {
            style: style.to_string(),
            width,
            color,
        })
    }

    // Legacy methods for backward compatibility - delegates to new methods
    #[allow(dead_code)]
    fn extract_paragraph(para: &docx_rs::Paragraph) -> StyledParagraph {
        // Create a minimal style resolver for backward compatibility
        // This is just a fallback; prefer using extract_paragraph_with_styles
        let mut styled = StyledParagraph::default();
        let props = &para.property;

        // Basic extraction without style resolution
        if let Some(ref align) = props.alignment {
            styled.alignment = format!("{:?}", align.val).to_lowercase();
        }

        if let Some(ref indent) = props.indent {
            if let Some(left) = indent.start {
                styled.left_indent = left as f32 * TWIPS_TO_POINTS;
            }
            if let Some(right) = indent.end {
                styled.right_indent = right as f32 * TWIPS_TO_POINTS;
            }
        }

        if let Some(ref style) = props.style {
            let style_id = &style.val;
            styled.style_id = Some(style_id.clone());
            if style_id.starts_with("Heading") || style_id.starts_with("heading") {
                styled.is_heading = true;
                styled.heading_level = style_id
                    .chars()
                    .last()
                    .and_then(|c| c.to_digit(10).map(|d| d as u8));
            }
        }

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

    #[allow(dead_code)]
    fn extract_run(run: &docx_rs::Run) -> StyledRun {
        let mut styled = StyledRun::default();

        for child in &run.children {
            match child {
                RunChild::Text(t) => styled.text.push_str(&t.text),
                RunChild::Tab(_) => styled.text.push('\t'),
                RunChild::Break(_) => styled.text.push('\n'),
                _ => {}
            }
        }

        let props = &run.run_property;
        styled.bold = props.bold.is_some();
        styled.italic = props.italic.is_some();
        styled.underline = props.underline.is_some();
        styled.strikethrough = props.strike.is_some();

        if let Some(ref sz) = props.sz {
            if let Ok(json) = serde_json::to_string(sz) {
                if let Ok(val) = json.trim_matches('"').parse::<f32>() {
                    styled.font_size = val * HALF_POINTS_TO_POINTS;
                }
            }
        }

        if let Some(ref color) = props.color {
            if let Ok(json) = serde_json::to_string(color) {
                if let Ok(parsed) = serde_json::from_str::<serde_json::Value>(&json) {
                    if let Some(val) = parsed.get("val").and_then(|v| v.as_str()) {
                        if val != "auto" {
                            styled.color = format!("#{}", val);
                        }
                    }
                }
            }
        }

        if let Some(ref fonts) = props.fonts {
            if let Ok(json) = serde_json::to_string(fonts) {
                if let Ok(parsed) = serde_json::from_str::<serde_json::Value>(&json) {
                    if let Some(ascii) = parsed.get("ascii").and_then(|v| v.as_str()) {
                        styled.font_family = ascii.to_string();
                    }
                }
            }
        }

        if let Some(ref highlight) = props.highlight {
            if let Ok(json) = serde_json::to_string(highlight) {
                let color_name = json.trim_matches('"').to_lowercase();
                styled.highlight = Some(Self::highlight_color_to_hex(&color_name));
            }
        }

        styled
    }

    #[allow(dead_code)]
    fn extract_table(table: &docx_rs::Table) -> Table {
        let mut result = Table::default();

        for width in &table.grid {
            result.column_widths.push(*width as f32 * TWIPS_TO_POINTS);
        }

        for table_child in &table.rows {
            let TableChild::TableRow(row) = table_child;
            let mut table_row = TableRow::default();

            for row_child in &row.cells {
                let TableRowChild::TableCell(cell) = row_child;
                let mut table_cell = TableCell::default();

                for content in &cell.children {
                    if let TableCellContent::Paragraph(para) = content {
                        table_cell.paragraphs.push(Self::extract_paragraph(para));
                    }
                }

                table_row.cells.push(table_cell);
            }
            result.rows.push(table_row);
        }

        result
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
    /// Now handles alignment, shading, borders, and list prefixes.
    fn layout_paragraph(
        &self,
        para: &StyledParagraph,
        x: f32,
        y: f32,
        width: f32,
        primitives: &mut Vec<crate::render_data::RenderPrimitive>,
    ) {
        use crate::render_data::{RenderPrimitive, RenderedLine, RenderedRect, RenderedText};
        use crate::renderer::Color;

        let left_margin = points_to_pixels(para.left_indent);
        let right_margin = points_to_pixels(para.right_indent);
        let first_line_offset = if para.hanging_indent > 0.0 {
            -points_to_pixels(para.hanging_indent)
        } else {
            points_to_pixels(para.first_line_indent)
        };

        let content_start_x = x + left_margin;
        let content_end_x = x + width - right_margin;
        let content_width = content_end_x - content_start_x;

        let line_height = para
            .runs
            .first()
            .map(|r| {
                let base = points_to_pixels(r.font_size);
                match para.line_spacing_rule.as_str() {
                    "exact" => points_to_pixels(para.line_spacing),
                    "atLeast" => points_to_pixels(para.line_spacing).max(base * 1.2),
                    _ => base * para.line_spacing * 1.2,
                }
            })
            .unwrap_or(20.0);

        let space_before = points_to_pixels(para.space_before);
        let space_after = points_to_pixels(para.space_after);

        // Estimate total paragraph height for shading/borders
        let total_text: String = para.runs.iter().map(|r| r.text.as_str()).collect();
        let avg_char_width = para
            .runs
            .first()
            .map(|r| points_to_pixels(r.font_size) * 0.5)
            .unwrap_or(6.0);
        let chars_per_line = (content_width / avg_char_width).max(1.0) as usize;
        let num_lines = (total_text.len() / chars_per_line).max(1) as f32;
        let para_height = num_lines * line_height + space_before + space_after;

        // Draw paragraph background/shading
        if let Some(ref shading) = para.shading {
            if let Some(color) = Color::from_hex(shading) {
                primitives.push(RenderPrimitive::Rect(RenderedRect {
                    x: content_start_x,
                    y,
                    width: content_width,
                    height: para_height,
                    fill: Some(color.to_rgba_array()),
                    stroke: None,
                }));
            }
        }

        // Draw paragraph borders
        let draw_border = |border: &ParagraphBorder,
                           x1: f32,
                           y1: f32,
                           x2: f32,
                           y2: f32,
                           primitives: &mut Vec<RenderPrimitive>| {
            if border.style != "none" {
                let color = Color::from_hex(&border.color)
                    .unwrap_or(Color::BLACK)
                    .to_rgba_array();
                let border_width = points_to_pixels(border.width).max(1.0);
                primitives.push(RenderPrimitive::Line(RenderedLine {
                    x1,
                    y1,
                    x2,
                    y2,
                    color,
                    width: border_width,
                }));
            }
        };

        if let Some(ref border) = para.border_top {
            draw_border(border, content_start_x, y, content_end_x, y, primitives);
        }
        if let Some(ref border) = para.border_bottom {
            draw_border(
                border,
                content_start_x,
                y + para_height,
                content_end_x,
                y + para_height,
                primitives,
            );
        }
        if let Some(ref border) = para.border_left {
            draw_border(
                border,
                content_start_x,
                y,
                content_start_x,
                y + para_height,
                primitives,
            );
        }
        if let Some(ref border) = para.border_right {
            draw_border(
                border,
                content_end_x,
                y,
                content_end_x,
                y + para_height,
                primitives,
            );
        }

        let mut current_y = y + space_before;
        let mut current_x = content_start_x + first_line_offset;
        let mut is_first_line = true;

        // Build list of words with their styles for proper layout
        struct WordInfo {
            text: String,
            font_family: String,
            font_size: f32,
            bold: bool,
            italic: bool,
            underline: bool,
            strikethrough: bool,
            color: [f32; 4],
            background: Option<[f32; 4]>,
            width: f32,
            vert_offset: f32,
        }

        let mut words: Vec<WordInfo> = Vec::new();

        // Add list prefix if present
        if let Some(ref prefix) = para.list_prefix {
            let first_run = para.runs.first();
            let font_family = first_run
                .map(|r| r.font_family.clone())
                .unwrap_or_else(|| "Calibri".to_string());
            let font_size = first_run.map(|r| r.font_size).unwrap_or(11.0);
            let char_width = points_to_pixels(font_size) * 0.5;
            let prefix_text = format!("{} ", prefix);

            words.push(WordInfo {
                width: prefix_text.len() as f32 * char_width,
                text: prefix_text,
                font_family,
                font_size,
                bold: first_run.map(|r| r.bold).unwrap_or(false),
                italic: first_run.map(|r| r.italic).unwrap_or(false),
                underline: false,
                strikethrough: false,
                color: [0.0, 0.0, 0.0, 1.0],
                background: None,
                vert_offset: 0.0,
            });
        }

        // Collect all words with their styles
        for run in &para.runs {
            // Skip hidden text
            if run.hidden {
                continue;
            }

            let style = run.to_text_style();
            let font_size = run.effective_font_size();
            let char_width = points_to_pixels(font_size) * 0.5;
            let bold_factor = if run.bold { 1.1 } else { 1.0 };
            let vert_offset = points_to_pixels(run.vertical_offset());

            // Handle text transformation (caps)
            let text = run.transformed_text();
            let text_words: Vec<&str> = text.split_whitespace().collect();

            for (i, word) in text_words.iter().enumerate() {
                let word_with_space = if i < text_words.len() - 1 {
                    format!("{} ", word)
                } else {
                    word.to_string()
                };

                let word_width = word_with_space.len() as f32 * char_width * bold_factor;

                words.push(WordInfo {
                    text: word_with_space,
                    font_family: run.font_family.clone(),
                    font_size,
                    bold: run.bold,
                    italic: run.italic,
                    underline: run.underline,
                    strikethrough: run.strikethrough || run.double_strikethrough,
                    color: style.color,
                    background: style.background,
                    width: word_width,
                    vert_offset,
                });
            }
        }

        // Layout words into lines
        struct LineInfo {
            words: Vec<(WordInfo, f32)>, // word and x position
            width: f32,
        }

        let mut lines: Vec<LineInfo> = Vec::new();
        let mut current_line = LineInfo {
            words: Vec::new(),
            width: 0.0,
        };
        let mut line_start_x = current_x;

        for word in words {
            let available_width = content_end_x - line_start_x;

            // Check if word fits on current line
            if current_line.width + word.width > available_width && !current_line.words.is_empty() {
                // Finish current line
                lines.push(current_line);
                current_line = LineInfo {
                    words: Vec::new(),
                    width: 0.0,
                };
                is_first_line = false;
                line_start_x = content_start_x; // No first line indent for subsequent lines
            }

            let word_x = current_line.width;
            current_line.width += word.width;
            current_line.words.push((word, word_x));
        }

        // Add remaining line
        if !current_line.words.is_empty() {
            lines.push(current_line);
        }

        // Render lines with proper alignment
        for (line_idx, line) in lines.iter().enumerate() {
            // Calculate line start position based on alignment
            let line_base_x = if line_idx == 0 {
                content_start_x + first_line_offset
            } else {
                content_start_x
            };

            let align_offset = match para.alignment.as_str() {
                "center" => (content_width - line.width) / 2.0,
                "right" => content_width - line.width,
                "both" | "justify" => {
                    // For justify, we'd need to add extra space between words
                    // For now, left-align the last line
                    if line_idx == lines.len() - 1 {
                        0.0
                    } else {
                        0.0
                    }
                }
                _ => 0.0, // left align
            };

            for (word, word_x) in &line.words {
                let render_x = line_base_x + word_x + align_offset;
                let render_y = current_y + line_height * 0.8 + word.vert_offset;

                primitives.push(RenderPrimitive::Text(RenderedText {
                    x: render_x,
                    y: render_y,
                    text: word.text.clone(),
                    font_family: word.font_family.clone(),
                    font_size: word.font_size,
                    bold: word.bold,
                    italic: word.italic,
                    underline: word.underline,
                    strikethrough: word.strikethrough,
                    color: word.color,
                    background: word.background,
                }));
            }

            current_y += line_height;
        }
    }

    /// Layout a table into render primitives without drawing.
    /// Now handles cell borders, backgrounds, vertical alignment, and proper cell sizing.
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

        // Calculate table position based on alignment
        let table_x = match table.alignment.as_str() {
            "center" => {
                let total_width: f32 = table
                    .column_widths
                    .iter()
                    .map(|w| points_to_pixels(*w))
                    .sum();
                x + (max_width - total_width.min(max_width)) / 2.0
            }
            "right" => {
                let total_width: f32 = table
                    .column_widths
                    .iter()
                    .map(|w| points_to_pixels(*w))
                    .sum();
                x + max_width - total_width.min(max_width)
            }
            _ => x + points_to_pixels(table.indent),
        };

        // Calculate scaled column widths
        let total_width: f32 = table
            .column_widths
            .iter()
            .map(|w| points_to_pixels(*w))
            .sum();
        let available_width = max_width - points_to_pixels(table.indent);
        let scale = if total_width > 0.0 && total_width > available_width {
            available_width / total_width
        } else {
            1.0
        };
        let scaled_widths: Vec<f32> = table
            .column_widths
            .iter()
            .map(|w| points_to_pixels(*w) * scale)
            .collect();

        let default_row_height = 25.0;
        let cell_padding = points_to_pixels(table.default_cell_margins.1); // Use right margin as horizontal padding

        let mut current_y = y;

        for (row_idx, row) in table.rows.iter().enumerate() {
            // Calculate row height
            let row_height = row
                .height
                .map(|h| points_to_pixels(h))
                .unwrap_or(default_row_height);

            let mut current_x = table_x;
            let mut col_position = 0usize;

            for cell in &row.cells {
                // Skip cells that are vertical merge continuations
                if cell.is_v_merge_continue {
                    // Still need to advance the column position
                    for _ in 0..cell.col_span {
                        current_x += scaled_widths.get(col_position).copied().unwrap_or(100.0);
                        col_position += 1;
                    }
                    continue;
                }

                // Calculate cell width (sum of spanned columns)
                let mut cell_width = 0.0;
                for span_idx in 0..cell.col_span as usize {
                    cell_width += scaled_widths
                        .get(col_position + span_idx)
                        .copied()
                        .unwrap_or(100.0);
                }

                // Draw cell background
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

                // Draw cell borders (use cell borders if defined, otherwise default black)
                let default_border_color = [0.0, 0.0, 0.0, 1.0];

                // Top border
                if let Some(ref border) = cell.border_top {
                    if border.style != "none" {
                        let border_color = Color::from_hex(&border.color)
                            .map(|c| c.to_rgba_array())
                            .unwrap_or(default_border_color);
                        let border_width = points_to_pixels(border.width).max(0.5);
                        primitives.push(RenderPrimitive::Line(RenderedLine {
                            x1: current_x,
                            y1: current_y,
                            x2: current_x + cell_width,
                            y2: current_y,
                            color: border_color,
                            width: border_width,
                        }));
                    }
                } else {
                    // Default border
                    primitives.push(RenderPrimitive::Line(RenderedLine {
                        x1: current_x,
                        y1: current_y,
                        x2: current_x + cell_width,
                        y2: current_y,
                        color: default_border_color,
                        width: 1.0,
                    }));
                }

                // Right border
                if let Some(ref border) = cell.border_right {
                    if border.style != "none" {
                        let border_color = Color::from_hex(&border.color)
                            .map(|c| c.to_rgba_array())
                            .unwrap_or(default_border_color);
                        let border_width = points_to_pixels(border.width).max(0.5);
                        primitives.push(RenderPrimitive::Line(RenderedLine {
                            x1: current_x + cell_width,
                            y1: current_y,
                            x2: current_x + cell_width,
                            y2: current_y + row_height,
                            color: border_color,
                            width: border_width,
                        }));
                    }
                } else {
                    primitives.push(RenderPrimitive::Line(RenderedLine {
                        x1: current_x + cell_width,
                        y1: current_y,
                        x2: current_x + cell_width,
                        y2: current_y + row_height,
                        color: default_border_color,
                        width: 1.0,
                    }));
                }

                // Bottom border
                if let Some(ref border) = cell.border_bottom {
                    if border.style != "none" {
                        let border_color = Color::from_hex(&border.color)
                            .map(|c| c.to_rgba_array())
                            .unwrap_or(default_border_color);
                        let border_width = points_to_pixels(border.width).max(0.5);
                        primitives.push(RenderPrimitive::Line(RenderedLine {
                            x1: current_x,
                            y1: current_y + row_height,
                            x2: current_x + cell_width,
                            y2: current_y + row_height,
                            color: border_color,
                            width: border_width,
                        }));
                    }
                } else {
                    primitives.push(RenderPrimitive::Line(RenderedLine {
                        x1: current_x,
                        y1: current_y + row_height,
                        x2: current_x + cell_width,
                        y2: current_y + row_height,
                        color: default_border_color,
                        width: 1.0,
                    }));
                }

                // Left border
                if let Some(ref border) = cell.border_left {
                    if border.style != "none" {
                        let border_color = Color::from_hex(&border.color)
                            .map(|c| c.to_rgba_array())
                            .unwrap_or(default_border_color);
                        let border_width = points_to_pixels(border.width).max(0.5);
                        primitives.push(RenderPrimitive::Line(RenderedLine {
                            x1: current_x,
                            y1: current_y,
                            x2: current_x,
                            y2: current_y + row_height,
                            color: border_color,
                            width: border_width,
                        }));
                    }
                } else {
                    primitives.push(RenderPrimitive::Line(RenderedLine {
                        x1: current_x,
                        y1: current_y,
                        x2: current_x,
                        y2: current_y + row_height,
                        color: default_border_color,
                        width: 1.0,
                    }));
                }

                // Render cell content (paragraphs)
                if !cell.paragraphs.is_empty() {
                    // Calculate text position based on vertical alignment
                    let cell_content_height = cell
                        .paragraphs
                        .iter()
                        .map(|p| {
                            p.runs
                                .first()
                                .map(|r| points_to_pixels(r.font_size) * 1.2)
                                .unwrap_or(20.0)
                        })
                        .sum::<f32>();

                    let text_start_y = match cell.vertical_align.as_str() {
                        "center" => current_y + (row_height - cell_content_height) / 2.0,
                        "bottom" => current_y + row_height - cell_content_height - cell.margins.2,
                        _ => current_y + cell.margins.0, // top
                    };

                    primitives.push(RenderPrimitive::Save);
                    primitives.push(RenderPrimitive::Clip {
                        x: current_x,
                        y: current_y,
                        width: cell_width,
                        height: row_height,
                    });

                    let mut para_y = text_start_y;
                    for para in &cell.paragraphs {
                        // Combine text from all runs
                        let text: String = para.runs.iter().map(|r| r.text.as_str()).collect();

                        if !text.is_empty() {
                            let style = para
                                .runs
                                .first()
                                .map(|r| r.to_text_style())
                                .unwrap_or_default();

                            let line_height = style.font_size * 1.2;

                            // Calculate x position based on paragraph alignment
                            let text_x = match para.alignment.as_str() {
                                "center" => {
                                    current_x + cell_width / 2.0
                                        - (text.len() as f32 * style.font_size * 0.25)
                                }
                                "right" => {
                                    current_x + cell_width
                                        - cell_padding
                                        - (text.len() as f32 * style.font_size * 0.5)
                                }
                                _ => current_x + cell_padding,
                            };

                            primitives.push(RenderPrimitive::Text(RenderedText {
                                x: text_x,
                                y: para_y + line_height * 0.8,
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

                            para_y += line_height;
                        }
                    }

                    primitives.push(RenderPrimitive::Restore);
                }

                // Advance to next cell
                for _ in 0..cell.col_span as usize {
                    current_x += scaled_widths.get(col_position).copied().unwrap_or(100.0);
                    col_position += 1;
                }
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
