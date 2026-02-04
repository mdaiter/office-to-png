//! Text layout and measurement for document rendering.

/// Simple 2D point.
#[derive(Clone, Copy, Debug)]
pub struct Point {
    pub x: f32,
    pub y: f32,
}

impl Point {
    pub fn new(x: f32, y: f32) -> Self {
        Self { x, y }
    }
}

/// Simple 2D rectangle.
#[derive(Clone, Copy, Debug)]
pub struct Rect {
    pub x: f32,
    pub y: f32,
    pub width: f32,
    pub height: f32,
}

impl Rect {
    pub fn new(x: f32, y: f32, width: f32, height: f32) -> Self {
        Self {
            x,
            y,
            width,
            height,
        }
    }

    pub fn contains(&self, point: Point) -> bool {
        point.x >= self.x
            && point.x <= self.x + self.width
            && point.y >= self.y
            && point.y <= self.y + self.height
    }
}

/// Text alignment options.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TextAlign {
    Left,
    Center,
    Right,
    Justify,
}

/// Text style information.
#[derive(Clone, Debug)]
pub struct TextStyle {
    /// Font family name
    pub font_family: String,
    /// Font size in points
    pub font_size: f32,
    /// Is bold
    pub bold: bool,
    /// Is italic
    pub italic: bool,
    /// Is underlined
    pub underline: bool,
    /// Is strikethrough
    pub strikethrough: bool,
    /// Text color (RGBA)
    pub color: [f32; 4],
    /// Background/highlight color (RGBA, None for transparent)
    pub background: Option<[f32; 4]>,
}

impl Default for TextStyle {
    fn default() -> Self {
        Self {
            font_family: "Arial".to_string(),
            font_size: 12.0,
            bold: false,
            italic: false,
            underline: false,
            strikethrough: false,
            color: [0.0, 0.0, 0.0, 1.0], // Black
            background: None,
        }
    }
}

impl TextStyle {
    /// Create a new text style with the given font.
    pub fn new(font_family: &str, font_size: f32) -> Self {
        Self {
            font_family: font_family.to_string(),
            font_size,
            ..Default::default()
        }
    }

    /// Set bold.
    pub fn bold(mut self, bold: bool) -> Self {
        self.bold = bold;
        self
    }

    /// Set italic.
    pub fn italic(mut self, italic: bool) -> Self {
        self.italic = italic;
        self
    }

    /// Set color.
    pub fn color(mut self, r: f32, g: f32, b: f32) -> Self {
        self.color = [r, g, b, 1.0];
        self
    }

    /// Get the CSS font string.
    pub fn to_css_font(&self) -> String {
        let style = if self.italic { "italic" } else { "normal" };
        let weight = if self.bold { "bold" } else { "normal" };
        format!(
            "{} {} {}px {}",
            style, weight, self.font_size, self.font_family
        )
    }
}

/// A run of text with consistent styling.
#[derive(Clone, Debug)]
pub struct TextRun {
    /// The text content
    pub text: String,
    /// Style for this run
    pub style: TextStyle,
}

impl TextRun {
    pub fn new(text: impl Into<String>, style: TextStyle) -> Self {
        Self {
            text: text.into(),
            style,
        }
    }

    /// Estimate the width of this text run in pixels.
    /// This is a rough estimate without proper font metrics.
    pub fn estimate_width(&self) -> f32 {
        let avg_char_width = self.style.font_size * 0.5; // Rough estimate
        let multiplier = if self.style.bold { 1.1 } else { 1.0 };
        self.text.len() as f32 * avg_char_width * multiplier
    }
}

/// A paragraph containing multiple text runs.
#[derive(Clone, Debug)]
pub struct Paragraph {
    /// Text runs in this paragraph
    pub runs: Vec<TextRun>,
    /// Text alignment
    pub align: TextAlign,
    /// Line spacing (1.0 = single, 1.5 = one and half, 2.0 = double)
    pub line_spacing: f32,
    /// Space before paragraph in points
    pub space_before: f32,
    /// Space after paragraph in points
    pub space_after: f32,
    /// First line indent in points
    pub first_line_indent: f32,
    /// Left indent in points
    pub left_indent: f32,
    /// Right indent in points
    pub right_indent: f32,
}

impl Default for Paragraph {
    fn default() -> Self {
        Self {
            runs: Vec::new(),
            align: TextAlign::Left,
            line_spacing: 1.0,
            space_before: 0.0,
            space_after: 0.0,
            first_line_indent: 0.0,
            left_indent: 0.0,
            right_indent: 0.0,
        }
    }
}

impl Paragraph {
    /// Create a new paragraph with the given runs.
    pub fn new(runs: Vec<TextRun>) -> Self {
        Self {
            runs,
            ..Default::default()
        }
    }

    /// Add a text run.
    pub fn add_run(&mut self, run: TextRun) {
        self.runs.push(run);
    }

    /// Get the combined text of all runs.
    pub fn text(&self) -> String {
        self.runs.iter().map(|r| r.text.as_str()).collect()
    }

    /// Estimate the height of this paragraph when laid out with the given width.
    pub fn estimate_height(&self, available_width: f32) -> f32 {
        if self.runs.is_empty() {
            return 0.0;
        }

        // Get the max font size for line height calculation
        let max_font_size = self
            .runs
            .iter()
            .map(|r| r.style.font_size)
            .fold(0.0f32, f32::max);

        let line_height = max_font_size * self.line_spacing;

        // Estimate number of lines
        let total_width: f32 = self.runs.iter().map(|r| r.estimate_width()).sum();
        let effective_width = available_width - self.left_indent - self.right_indent;
        let num_lines = (total_width / effective_width).ceil().max(1.0);

        self.space_before + (num_lines * line_height) + self.space_after
    }
}

/// A laid-out line of text, ready for rendering.
#[derive(Clone, Debug)]
pub struct LayoutLine {
    /// Position of the line
    pub position: Point,
    /// Text runs on this line (may be partial runs)
    pub runs: Vec<(TextRun, f32)>, // (run, x_offset)
    /// Line height
    pub height: f32,
    /// Baseline offset from top
    pub baseline: f32,
}

/// Result of laying out a paragraph.
#[derive(Clone, Debug)]
pub struct LayoutResult {
    /// Laid out lines
    pub lines: Vec<LayoutLine>,
    /// Total height used
    pub total_height: f32,
    /// Bounding rect
    pub bounds: Rect,
}

/// Simple text layout engine.
pub struct TextLayoutEngine {
    /// Default font family
    pub default_font: String,
    /// Default font size
    pub default_size: f32,
}

impl Default for TextLayoutEngine {
    fn default() -> Self {
        Self {
            default_font: "Arial".to_string(),
            default_size: 12.0,
        }
    }
}

impl TextLayoutEngine {
    /// Create a new layout engine.
    pub fn new() -> Self {
        Self::default()
    }

    /// Layout a paragraph within the given bounds.
    pub fn layout_paragraph(&self, paragraph: &Paragraph, bounds: Rect) -> LayoutResult {
        let mut lines = Vec::new();
        let mut y = bounds.y + paragraph.space_before;

        if paragraph.runs.is_empty() {
            return LayoutResult {
                lines,
                total_height: paragraph.space_before + paragraph.space_after,
                bounds: Rect::new(
                    bounds.x,
                    bounds.y,
                    bounds.width,
                    paragraph.space_before + paragraph.space_after,
                ),
            };
        }

        // Get max font size for line height
        let max_font_size = paragraph
            .runs
            .iter()
            .map(|r| r.style.font_size)
            .fold(self.default_size, f32::max);

        let line_height = max_font_size * paragraph.line_spacing;
        let baseline = max_font_size * 0.8; // Rough baseline estimate

        let effective_width = bounds.width - paragraph.left_indent - paragraph.right_indent;

        // Simple word wrapping
        let mut current_line_runs: Vec<(TextRun, f32)> = Vec::new();
        let mut current_x = paragraph.first_line_indent;
        let mut is_first_line = true;

        for run in &paragraph.runs {
            let words: Vec<&str> = run.text.split_whitespace().collect();

            for (i, word) in words.iter().enumerate() {
                let word_with_space = if i < words.len() - 1 {
                    format!("{} ", word)
                } else {
                    word.to_string()
                };

                let word_run = TextRun::new(&word_with_space, run.style.clone());
                let word_width = word_run.estimate_width();

                // Check if word fits on current line
                if current_x + word_width > effective_width && !current_line_runs.is_empty() {
                    // Start new line
                    let line_x = match paragraph.align {
                        TextAlign::Left | TextAlign::Justify => bounds.x + paragraph.left_indent,
                        TextAlign::Right => {
                            bounds.x + bounds.width - paragraph.right_indent - current_x
                        }
                        TextAlign::Center => bounds.x + (bounds.width - current_x) / 2.0,
                    };

                    lines.push(LayoutLine {
                        position: Point::new(line_x, y),
                        runs: std::mem::take(&mut current_line_runs),
                        height: line_height,
                        baseline,
                    });

                    y += line_height;
                    current_x = if is_first_line { 0.0 } else { 0.0 };
                    is_first_line = false;
                }

                current_line_runs.push((word_run, current_x));
                current_x += word_width;
            }
        }

        // Add remaining text
        if !current_line_runs.is_empty() {
            let line_x = match paragraph.align {
                TextAlign::Left | TextAlign::Justify => bounds.x + paragraph.left_indent,
                TextAlign::Right => bounds.x + bounds.width - paragraph.right_indent - current_x,
                TextAlign::Center => bounds.x + (bounds.width - current_x) / 2.0,
            };

            lines.push(LayoutLine {
                position: Point::new(line_x, y),
                runs: current_line_runs,
                height: line_height,
                baseline,
            });

            y += line_height;
        }

        let total_height = y - bounds.y + paragraph.space_after;

        LayoutResult {
            lines,
            total_height,
            bounds: Rect::new(bounds.x, bounds.y, bounds.width, total_height),
        }
    }
}
