//! WebGPU render context for document rendering.

use wasm_bindgen::prelude::*;
use wasm_bindgen_futures::JsFuture;
use web_sys::{HtmlCanvasElement, GpuCanvasContext, GpuDevice, GpuQueue, GpuTexture};
use std::sync::Arc;

/// WebGPU rendering context.
pub struct RenderContext {
    /// The canvas element
    canvas: HtmlCanvasElement,
    /// WebGPU device
    device: GpuDevice,
    /// WebGPU queue
    queue: GpuQueue,
    /// Canvas context
    context: GpuCanvasContext,
    /// Current texture format
    format: web_sys::GpuTextureFormat,
    /// Canvas width
    width: u32,
    /// Canvas height
    height: u32,
    /// Current zoom level
    zoom: f32,
    /// Background color (RGBA)
    background_color: [f32; 4],
}

impl RenderContext {
    /// Create a new render context for a canvas.
    pub async fn new(canvas: HtmlCanvasElement) -> Result<Self, JsValue> {
        let window = web_sys::window().ok_or("No window")?;
        let navigator = window.navigator();
        let gpu = navigator.gpu();
        
        if gpu.is_undefined() {
            return Err(JsValue::from_str("WebGPU not supported in this browser"));
        }
        
        let gpu: web_sys::Gpu = gpu.unchecked_into();
        
        // Request adapter
        let adapter_promise = gpu.request_adapter();
        let adapter = JsFuture::from(adapter_promise).await?;
        
        if adapter.is_null() || adapter.is_undefined() {
            return Err(JsValue::from_str("Failed to get WebGPU adapter"));
        }
        
        let adapter: web_sys::GpuAdapter = adapter.unchecked_into();
        
        // Request device
        let device_promise = adapter.request_device();
        let device = JsFuture::from(device_promise).await?;
        let device: GpuDevice = device.unchecked_into();
        let queue = device.queue();
        
        // Get canvas context
        let context = canvas
            .get_context("webgpu")?
            .ok_or("Failed to get WebGPU context")?
            .dyn_into::<GpuCanvasContext>()?;
        
        // Configure canvas
        let format = gpu.get_preferred_canvas_format();
        let config = web_sys::GpuCanvasConfiguration::new(&device, format);
        context.configure(&config);
        
        let width = canvas.width();
        let height = canvas.height();
        
        Ok(Self {
            canvas,
            device,
            queue,
            context,
            format,
            width,
            height,
            zoom: 1.0,
            background_color: [1.0, 1.0, 1.0, 1.0], // White
        })
    }
    
    /// Resize the render context.
    pub fn resize(&mut self, width: u32, height: u32) -> Result<(), JsValue> {
        self.width = width;
        self.height = height;
        self.canvas.set_width(width);
        self.canvas.set_height(height);
        
        // Reconfigure context
        let config = web_sys::GpuCanvasConfiguration::new(&self.device, self.format);
        self.context.configure(&config);
        
        Ok(())
    }
    
    /// Get canvas width.
    pub fn width(&self) -> u32 {
        self.width
    }
    
    /// Get canvas height.
    pub fn height(&self) -> u32 {
        self.height
    }
    
    /// Set zoom level.
    pub fn set_zoom(&mut self, zoom: f32) {
        self.zoom = zoom;
    }
    
    /// Get zoom level.
    pub fn zoom(&self) -> f32 {
        self.zoom
    }
    
    /// Set background color.
    pub fn set_background_color(&mut self, r: f32, g: f32, b: f32, a: f32) {
        self.background_color = [r, g, b, a];
    }
    
    /// Clear the canvas with the background color.
    pub fn clear(&self) -> Result<(), JsValue> {
        let texture = self.context.get_current_texture();
        let view = texture.create_view();
        
        let encoder = self.device.create_command_encoder();
        
        // Create color attachment
        let color_attachment = web_sys::GpuRenderPassColorAttachment::new(
            web_sys::GpuLoadOp::Clear,
            web_sys::GpuStoreOp::Store,
            &view,
        );
        
        // Set clear color
        let clear_value = js_sys::Object::new();
        js_sys::Reflect::set(&clear_value, &"r".into(), &self.background_color[0].into())?;
        js_sys::Reflect::set(&clear_value, &"g".into(), &self.background_color[1].into())?;
        js_sys::Reflect::set(&clear_value, &"b".into(), &self.background_color[2].into())?;
        js_sys::Reflect::set(&clear_value, &"a".into(), &self.background_color[3].into())?;
        color_attachment.set_clear_value(&clear_value);
        
        // Create render pass
        let color_attachments = js_sys::Array::new();
        color_attachments.push(&color_attachment);
        
        let render_pass_desc = web_sys::GpuRenderPassDescriptor::new(&color_attachments);
        let pass = encoder.begin_render_pass(&render_pass_desc);
        pass.end();
        
        // Submit
        let command_buffer = encoder.finish();
        let commands = js_sys::Array::new();
        commands.push(&command_buffer);
        self.queue.submit(&commands);
        
        Ok(())
    }
    
    /// Draw a filled rectangle.
    pub fn draw_rect(&self, x: f32, y: f32, width: f32, height: f32, color: [f32; 4]) -> Result<(), JsValue> {
        // For now, we'll use 2D canvas fallback for simple shapes
        // A full WebGPU implementation would use vertex/fragment shaders
        
        // This is a simplified implementation - a production version would
        // use proper WebGPU rendering pipelines
        
        Ok(())
    }
    
    /// Draw text at the specified position.
    pub fn draw_text(&self, text: &str, x: f32, y: f32, font_size: f32, color: [f32; 4]) -> Result<(), JsValue> {
        // Text rendering in WebGPU is complex - would need:
        // 1. Font atlas generation
        // 2. Glyph rendering with signed distance fields
        // 3. Or use Canvas2D for text and composite
        
        // For now, this is a placeholder
        Ok(())
    }
    
    /// Export the current canvas content to PNG bytes.
    pub async fn export_png(&self) -> Result<Vec<u8>, JsValue> {
        // Get canvas data URL and convert to bytes
        let data_url = self.canvas.to_data_url_with_type("image/png")?;
        
        // Parse data URL (format: data:image/png;base64,...)
        let base64_data = data_url
            .strip_prefix("data:image/png;base64,")
            .ok_or("Invalid data URL format")?;
        
        // Decode base64
        let decoded = base64::decode(base64_data)
            .map_err(|e| JsValue::from_str(&format!("Base64 decode error: {}", e)))?;
        
        Ok(decoded)
    }
    
    /// Get the WebGPU device.
    pub fn device(&self) -> &GpuDevice {
        &self.device
    }
    
    /// Get the WebGPU queue.
    pub fn queue(&self) -> &GpuQueue {
        &self.queue
    }
    
    /// Get the current texture.
    pub fn current_texture(&self) -> GpuTexture {
        self.context.get_current_texture()
    }
}

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
        Self { x, y, width, height }
    }
    
    pub fn contains(&self, point: Point) -> bool {
        point.x >= self.x 
            && point.x <= self.x + self.width
            && point.y >= self.y
            && point.y <= self.y + self.height
    }
}

// Add base64 dependency for PNG export
mod base64 {
    const ALPHABET: &[u8] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    
    pub fn decode(input: &str) -> Result<Vec<u8>, &'static str> {
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
                _ => return Err("Invalid base64 character"),
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
}
