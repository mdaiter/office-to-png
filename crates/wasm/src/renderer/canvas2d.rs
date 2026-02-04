//! Canvas 2D rendering backend.
//!
//! This is the primary rendering backend that uses the HTML5 Canvas 2D API.
//! It provides good compatibility across all browsers.

use super::traits::{BorderStyle, Color, RenderBackend, TextMetrics};
use crate::text_layout::TextStyle;
use wasm_bindgen::prelude::*;
use wasm_bindgen::Clamped;
use web_sys::{CanvasRenderingContext2d, HtmlCanvasElement, ImageData};

/// Canvas 2D rendering backend.
pub struct Canvas2DRenderer {
    canvas: HtmlCanvasElement,
    ctx: CanvasRenderingContext2d,
    width: f32,
    height: f32,
    /// Device pixel ratio for high-DPI displays
    dpr: f32,
}

impl Canvas2DRenderer {
    /// Create a new Canvas 2D renderer for the given canvas ID.
    pub fn new(canvas_id: &str) -> Result<Self, String> {
        let window = web_sys::window().ok_or("No window")?;
        let document = window.document().ok_or("No document")?;

        let element = document
            .get_element_by_id(canvas_id)
            .ok_or_else(|| format!("Canvas element '{}' not found", canvas_id))?;

        let canvas: HtmlCanvasElement =
            element.dyn_into().map_err(|_| "Element is not a canvas")?;

        let ctx = canvas
            .get_context("2d")
            .map_err(|e| format!("Failed to get 2d context: {:?}", e))?
            .ok_or("No 2d context")?
            .dyn_into::<CanvasRenderingContext2d>()
            .map_err(|_| "Failed to cast to CanvasRenderingContext2d")?;

        // Get device pixel ratio for high-DPI displays
        let dpr = window.device_pixel_ratio() as f32;

        let width = canvas.client_width() as f32;
        let height = canvas.client_height() as f32;

        let renderer = Self {
            canvas,
            ctx,
            width,
            height,
            dpr,
        };

        // Set up high-DPI scaling
        renderer.setup_high_dpi()?;

        Ok(renderer)
    }

    /// Create from an existing canvas element.
    pub fn from_canvas(canvas: HtmlCanvasElement) -> Result<Self, String> {
        let ctx = canvas
            .get_context("2d")
            .map_err(|e| format!("Failed to get 2d context: {:?}", e))?
            .ok_or("No 2d context")?
            .dyn_into::<CanvasRenderingContext2d>()
            .map_err(|_| "Failed to cast to CanvasRenderingContext2d")?;

        let window = web_sys::window().ok_or("No window")?;
        let dpr = window.device_pixel_ratio() as f32;

        let width = canvas.client_width() as f32;
        let height = canvas.client_height() as f32;

        let renderer = Self {
            canvas,
            ctx,
            width,
            height,
            dpr,
        };

        renderer.setup_high_dpi()?;

        Ok(renderer)
    }

    /// Set up high-DPI rendering
    fn setup_high_dpi(&self) -> Result<(), String> {
        // Scale canvas buffer size for high-DPI
        let buffer_width = (self.width * self.dpr) as u32;
        let buffer_height = (self.height * self.dpr) as u32;

        self.canvas.set_width(buffer_width);
        self.canvas.set_height(buffer_height);

        // Scale the context so drawing operations work in CSS pixels
        self.ctx
            .scale(self.dpr as f64, self.dpr as f64)
            .map_err(|e| format!("Failed to scale context: {:?}", e))?;

        Ok(())
    }

    /// Get the underlying canvas element.
    pub fn canvas(&self) -> &HtmlCanvasElement {
        &self.canvas
    }

    /// Get the 2D rendering context.
    pub fn context(&self) -> &CanvasRenderingContext2d {
        &self.ctx
    }

    /// Get the device pixel ratio.
    pub fn device_pixel_ratio(&self) -> f32 {
        self.dpr
    }

    /// Set font on the context.
    fn set_font(&self, style: &TextStyle) {
        self.ctx.set_font(&style.to_css_font());
    }

    /// Set fill color.
    fn set_fill_color(&self, color: Color) {
        self.ctx.set_fill_style_str(&color.to_css());
    }

    /// Set stroke color.
    fn set_stroke_color(&self, color: Color) {
        self.ctx.set_stroke_style_str(&color.to_css());
    }

    /// Draw underline for text.
    fn draw_underline(&self, x: f32, y: f32, width: f32, style: &TextStyle) {
        let line_y = y + style.font_size * 0.1; // Slightly below baseline
        self.ctx.begin_path();
        self.ctx.move_to(x as f64, line_y as f64);
        self.ctx.line_to((x + width) as f64, line_y as f64);
        self.set_stroke_color(Color::from_rgba_array(style.color));
        self.ctx.set_line_width(1.0);
        self.ctx.stroke();
    }

    /// Draw strikethrough for text.
    fn draw_strikethrough(&self, x: f32, y: f32, width: f32, style: &TextStyle) {
        let line_y = y - style.font_size * 0.3; // Middle of text
        self.ctx.begin_path();
        self.ctx.move_to(x as f64, line_y as f64);
        self.ctx.line_to((x + width) as f64, line_y as f64);
        self.set_stroke_color(Color::from_rgba_array(style.color));
        self.ctx.set_line_width(1.0);
        self.ctx.stroke();
    }
}

impl RenderBackend for Canvas2DRenderer {
    fn width(&self) -> f32 {
        self.width
    }

    fn height(&self) -> f32 {
        self.height
    }

    fn resize(&mut self, width: f32, height: f32) -> Result<(), String> {
        self.width = width;
        self.height = height;

        // Update canvas buffer size
        let buffer_width = (width * self.dpr) as u32;
        let buffer_height = (height * self.dpr) as u32;

        self.canvas.set_width(buffer_width);
        self.canvas.set_height(buffer_height);

        // Reset scale after resize
        self.ctx
            .set_transform(1.0, 0.0, 0.0, 1.0, 0.0, 0.0)
            .map_err(|e| format!("Failed to reset transform: {:?}", e))?;
        self.ctx
            .scale(self.dpr as f64, self.dpr as f64)
            .map_err(|e| format!("Failed to scale context: {:?}", e))?;

        Ok(())
    }

    fn clear(&self, color: Color) -> Result<(), String> {
        self.set_fill_color(color);
        self.ctx
            .fill_rect(0.0, 0.0, self.width as f64, self.height as f64);
        Ok(())
    }

    fn fill_rect(
        &self,
        x: f32,
        y: f32,
        width: f32,
        height: f32,
        color: Color,
    ) -> Result<(), String> {
        self.set_fill_color(color);
        self.ctx
            .fill_rect(x as f64, y as f64, width as f64, height as f64);
        Ok(())
    }

    fn stroke_rect(
        &self,
        x: f32,
        y: f32,
        width: f32,
        height: f32,
        border: &BorderStyle,
    ) -> Result<(), String> {
        self.set_stroke_color(border.color);
        self.ctx.set_line_width(border.width as f64);

        // Handle dash pattern
        if let Some(ref pattern) = border.dash_pattern {
            let js_pattern = js_sys::Array::new();
            for &v in pattern {
                js_pattern.push(&JsValue::from_f64(v as f64));
            }
            self.ctx
                .set_line_dash(&js_pattern)
                .map_err(|e| format!("Failed to set line dash: {:?}", e))?;
        } else {
            self.ctx
                .set_line_dash(&js_sys::Array::new())
                .map_err(|e| format!("Failed to clear line dash: {:?}", e))?;
        }

        self.ctx
            .stroke_rect(x as f64, y as f64, width as f64, height as f64);
        Ok(())
    }

    fn draw_text(&self, text: &str, x: f32, y: f32, style: &TextStyle) -> Result<(), String> {
        // Draw background if present
        if let Some(bg) = style.background {
            let metrics = self.measure_text(text, style)?;
            self.set_fill_color(Color::from_rgba_array(bg));
            self.ctx.fill_rect(
                x as f64,
                (y - metrics.ascent) as f64,
                metrics.width as f64,
                metrics.height as f64,
            );
        }

        // Set font and color
        self.set_font(style);
        self.set_fill_color(Color::from_rgba_array(style.color));

        // Draw the text
        self.ctx
            .fill_text(text, x as f64, y as f64)
            .map_err(|e| format!("Failed to draw text: {:?}", e))?;

        // Draw decorations
        let text_width = self.measure_text(text, style)?.width;

        if style.underline {
            self.draw_underline(x, y, text_width, style);
        }

        if style.strikethrough {
            self.draw_strikethrough(x, y, text_width, style);
        }

        Ok(())
    }

    fn measure_text(&self, text: &str, style: &TextStyle) -> Result<TextMetrics, String> {
        self.set_font(style);

        let metrics = self
            .ctx
            .measure_text(text)
            .map_err(|e| format!("Failed to measure text: {:?}", e))?;

        let width = metrics.width() as f32;

        // Canvas 2D TextMetrics has limited information
        // We estimate height based on font size
        let height = style.font_size;
        let ascent = style.font_size * 0.8;
        let descent = style.font_size * 0.2;

        Ok(TextMetrics {
            width,
            height,
            ascent,
            descent,
        })
    }

    fn draw_line(
        &self,
        x1: f32,
        y1: f32,
        x2: f32,
        y2: f32,
        color: Color,
        width: f32,
    ) -> Result<(), String> {
        self.ctx.begin_path();
        self.ctx.move_to(x1 as f64, y1 as f64);
        self.ctx.line_to(x2 as f64, y2 as f64);
        self.set_stroke_color(color);
        self.ctx.set_line_width(width as f64);
        self.ctx.stroke();
        Ok(())
    }

    fn draw_image(
        &self,
        data: &[u8],
        img_width: u32,
        img_height: u32,
        x: f32,
        y: f32,
        dest_width: f32,
        dest_height: f32,
    ) -> Result<(), String> {
        // Create ImageData from raw RGBA bytes
        let clamped = Clamped(data);
        let image_data =
            ImageData::new_with_u8_clamped_array_and_sh(clamped, img_width, img_height)
                .map_err(|e| format!("Failed to create ImageData: {:?}", e))?;

        // For scaling, we need to use a temporary canvas
        // First draw at original size, then draw scaled
        if (dest_width - img_width as f32).abs() < 0.1
            && (dest_height - img_height as f32).abs() < 0.1
        {
            // No scaling needed
            self.ctx
                .put_image_data(&image_data, x as f64, y as f64)
                .map_err(|e| format!("Failed to put image data: {:?}", e))?;
        } else {
            // Create a temporary canvas for scaling
            let document = web_sys::window()
                .ok_or("No window")?
                .document()
                .ok_or("No document")?;

            let temp_canvas: HtmlCanvasElement = document
                .create_element("canvas")
                .map_err(|e| format!("Failed to create temp canvas: {:?}", e))?
                .dyn_into()
                .map_err(|_| "Failed to cast to canvas")?;

            temp_canvas.set_width(img_width);
            temp_canvas.set_height(img_height);

            let temp_ctx = temp_canvas
                .get_context("2d")
                .map_err(|e| format!("Failed to get temp context: {:?}", e))?
                .ok_or("No temp context")?
                .dyn_into::<CanvasRenderingContext2d>()
                .map_err(|_| "Failed to cast temp context")?;

            temp_ctx
                .put_image_data(&image_data, 0.0, 0.0)
                .map_err(|e| format!("Failed to put temp image data: {:?}", e))?;

            // Draw scaled to main canvas
            self.ctx
                .draw_image_with_html_canvas_element_and_dw_and_dh(
                    &temp_canvas,
                    x as f64,
                    y as f64,
                    dest_width as f64,
                    dest_height as f64,
                )
                .map_err(|e| format!("Failed to draw scaled image: {:?}", e))?;
        }

        Ok(())
    }

    fn save(&self) -> Result<(), String> {
        self.ctx.save();
        Ok(())
    }

    fn restore(&self) -> Result<(), String> {
        self.ctx.restore();
        Ok(())
    }

    fn translate(&self, x: f32, y: f32) -> Result<(), String> {
        self.ctx
            .translate(x as f64, y as f64)
            .map_err(|e| format!("Failed to translate: {:?}", e))?;
        Ok(())
    }

    fn scale(&self, x: f32, y: f32) -> Result<(), String> {
        self.ctx
            .scale(x as f64, y as f64)
            .map_err(|e| format!("Failed to scale: {:?}", e))?;
        Ok(())
    }

    fn clip(&self, x: f32, y: f32, width: f32, height: f32) -> Result<(), String> {
        self.ctx.begin_path();
        self.ctx
            .rect(x as f64, y as f64, width as f64, height as f64);
        self.ctx.clip();
        Ok(())
    }

    fn export_png(&self) -> Result<Vec<u8>, String> {
        let data_url = self
            .canvas
            .to_data_url_with_type("image/png")
            .map_err(|e| format!("Failed to get data URL: {:?}", e))?;

        // Parse data URL (format: data:image/png;base64,...)
        let base64_data = data_url
            .strip_prefix("data:image/png;base64,")
            .ok_or("Invalid data URL format")?;

        // Decode base64
        decode_base64(base64_data)
    }
}

/// Simple base64 decoder
fn decode_base64(input: &str) -> Result<Vec<u8>, String> {
    let input = input.as_bytes();
    let mut output = Vec::with_capacity(input.len() * 3 / 4);

    let mut buffer = 0u32;
    let mut bits = 0;

    for &byte in input {
        if byte == b'=' {
            break;
        }

        let value = match byte {
            b'A'..=b'Z' => byte - b'A',
            b'a'..=b'z' => byte - b'a' + 26,
            b'0'..=b'9' => byte - b'0' + 52,
            b'+' => 62,
            b'/' => 63,
            b'\n' | b'\r' | b' ' | b'\t' => continue,
            _ => return Err("Invalid base64 character".to_string()),
        };

        buffer = (buffer << 6) | value as u32;
        bits += 6;

        if bits >= 8 {
            bits -= 8;
            output.push((buffer >> bits) as u8);
            buffer &= (1 << bits) - 1;
        }
    }

    Ok(output)
}
