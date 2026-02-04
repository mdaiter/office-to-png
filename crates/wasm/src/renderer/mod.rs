//! Rendering backends for document visualization.
//!
//! This module provides different rendering backends:
//! - Canvas 2D: Primary backend using HTML5 Canvas 2D API
//! - WebGPU: High-performance GPU-accelerated backend

mod canvas2d;
mod traits;
#[cfg(feature = "webgpu")]
mod webgpu;

pub use canvas2d::Canvas2DRenderer;
pub use traits::{BorderStyle, Color, DrawCommand, RenderBackend, TextMetrics};
#[cfg(feature = "webgpu")]
pub use webgpu::WebGPURenderer;

/// Default page dimensions (US Letter at 96 DPI)
pub const DEFAULT_PAGE_WIDTH: f32 = 816.0; // 8.5 inches * 96
pub const DEFAULT_PAGE_HEIGHT: f32 = 1056.0; // 11 inches * 96

/// Default margins in pixels (at 96 DPI)
pub const DEFAULT_MARGIN: f32 = 72.0; // 0.75 inches

/// Points per inch
pub const POINTS_PER_INCH: f32 = 72.0;

/// Pixels per inch (screen resolution)
pub const PIXELS_PER_INCH: f32 = 96.0;

/// Convert points to pixels
pub fn points_to_pixels(points: f32) -> f32 {
    points * PIXELS_PER_INCH / POINTS_PER_INCH
}

/// Convert pixels to points
pub fn pixels_to_points(pixels: f32) -> f32 {
    pixels * POINTS_PER_INCH / PIXELS_PER_INCH
}

/// Convert EMUs (English Metric Units) to pixels
/// 1 inch = 914400 EMUs, 1 inch = 96 pixels at screen DPI
pub fn emu_to_pixels(emu: i64) -> f32 {
    (emu as f32) * PIXELS_PER_INCH / 914400.0
}

/// Convert twips to pixels (1 twip = 1/20 point)
pub fn twips_to_pixels(twips: i32) -> f32 {
    points_to_pixels(twips as f32 / 20.0)
}
