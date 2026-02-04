//! PDF to PNG rendering using pdfium (Google's PDF engine).
//!
//! This module provides high-performance PDF rendering with:
//! - Parallel page rendering via rayon
//! - Configurable DPI output
//! - SIMD-accelerated PNG encoding (via zlib-rs)

use crate::config::{PngPage, RenderConfig};
use crate::error::{ConversionError, Result};
use image::{ImageBuffer, Rgba, RgbaImage};
use pdfium_render::prelude::*;
use rayon::iter::IntoParallelIterator;
use rayon::prelude::*;
use std::io::Cursor;
use std::path::{Path, PathBuf};
use std::sync::Arc;
use std::time::Instant;
use tracing::{debug, error, info, warn};

/// PDF to PNG renderer using pdfium.
pub struct PdfRenderer {
    /// Render configuration.
    config: RenderConfig,
    /// Pdfium library instance.
    pdfium: Arc<Pdfium>,
    /// Rayon thread pool for parallel rendering.
    thread_pool: rayon::ThreadPool,
}

impl PdfRenderer {
    /// Create a new PDF renderer.
    pub fn new(config: RenderConfig) -> Result<Self> {
        config.validate()?;

        // Initialize pdfium
        // Try to bind to system library first, then fall back to bundled
        let pdfium = Pdfium::new(
            Pdfium::bind_to_library(Pdfium::pdfium_platform_library_name_at_path("./"))
                .or_else(|_| {
                    Pdfium::bind_to_library(Pdfium::pdfium_platform_library_name_at_path(
                        "/usr/lib",
                    ))
                })
                .or_else(|_| {
                    Pdfium::bind_to_library(Pdfium::pdfium_platform_library_name_at_path(
                        "/usr/local/lib",
                    ))
                })
                .or_else(|_| Pdfium::bind_to_system_library())
                .map_err(|e| {
                    ConversionError::PdfiumError(format!("Failed to load pdfium library: {}", e))
                })?,
        );

        // Create thread pool
        let thread_pool = rayon::ThreadPoolBuilder::new()
            .num_threads(config.render_threads)
            .build()
            .map_err(|e| {
                ConversionError::InvalidConfig(format!("Failed to create thread pool: {}", e))
            })?;

        info!(
            "PDF renderer initialized with {} threads, {} DPI",
            config.render_threads, config.dpi
        );

        Ok(Self {
            config,
            pdfium: Arc::new(pdfium),
            thread_pool,
        })
    }

    /// Get the configured DPI.
    pub fn dpi(&self) -> u32 {
        self.config.dpi
    }

    /// Render all pages of a PDF to PNG images.
    ///
    /// Note: pdfium's PdfDocument is not thread-safe, so pages are rendered sequentially.
    /// However, PNG encoding is parallelized using rayon.
    pub fn render_all_pages(&self, pdf_path: &Path) -> Result<Vec<PngPage>> {
        let start = Instant::now();

        // Load the PDF
        let document = self
            .pdfium
            .load_pdf_from_file(pdf_path, None)
            .map_err(|e| ConversionError::PdfRenderError(format!("Failed to load PDF: {}", e)))?;

        let page_count = document.pages().len() as usize;
        debug!("Rendering {} pages from {:?}", page_count, pdf_path);

        if page_count == 0 {
            return Ok(vec![]);
        }

        // Render pages sequentially (pdfium is not thread-safe)
        // We collect raw image data first, then parallelize PNG encoding
        let mut raw_images: Vec<(usize, RgbaImage)> = Vec::with_capacity(page_count);

        for page_idx in 0..page_count {
            let page = document.pages().get(page_idx as u16).map_err(|e| {
                ConversionError::PdfRenderError(format!(
                    "Failed to get page {}: {}",
                    page_idx + 1,
                    e
                ))
            })?;

            // Calculate pixel dimensions based on DPI
            let scale = self.config.dpi as f32 / 72.0;
            let width = (page.width().value * scale) as u32;
            let height = (page.height().value * scale) as u32;

            // Create render config
            let render_config = PdfRenderConfig::new()
                .set_target_width(width as i32)
                .set_target_height(height as i32)
                .rotate_if_landscape(PdfPageRenderRotation::None, false);

            // Render to bitmap
            let bitmap = page.render_with_config(&render_config).map_err(|e| {
                ConversionError::PdfRenderError(format!(
                    "Failed to render page {}: {}",
                    page_idx + 1,
                    e
                ))
            })?;

            // Convert to image buffer
            let image_result = bitmap.as_image();
            let rgba_image: RgbaImage = image_result.into_rgba8();

            // Apply background color if not using alpha
            let final_image = if !self.config.use_alpha {
                self.apply_background(rgba_image)
            } else {
                rgba_image
            };

            raw_images.push((page_idx, final_image));
        }

        // Now parallelize PNG encoding (using standalone function to avoid Send issues)
        let rendered_pages: Vec<Result<PngPage>> = self.thread_pool.install(|| {
            raw_images
                .into_par_iter()
                .map(|(page_idx, image)| {
                    let png_data = encode_png_standalone(&image)?;
                    Ok(PngPage {
                        page_number: page_idx + 1,
                        data: png_data,
                        width: image.width(),
                        height: image.height(),
                        output_path: None,
                    })
                })
                .collect()
        });

        // Collect results, handling errors
        let mut pages = Vec::with_capacity(page_count);
        for result in rendered_pages {
            match result {
                Ok(page) => pages.push(page),
                Err(e) => {
                    error!("Failed to encode page: {:?}", e);
                    return Err(e);
                }
            }
        }

        // Sort by page number
        pages.sort_by_key(|p| p.page_number);

        debug!("Rendered {} pages in {:?}", page_count, start.elapsed());

        Ok(pages)
    }

    /// Render a single page to PNG.
    fn render_single_page(&self, document: &PdfDocument, page_idx: usize) -> Result<PngPage> {
        let page = document.pages().get(page_idx as u16).map_err(|e| {
            ConversionError::PdfRenderError(format!("Failed to get page {}: {}", page_idx + 1, e))
        })?;

        // Calculate pixel dimensions based on DPI
        // PDF pages are typically 72 DPI, so we scale up
        let scale = self.config.dpi as f32 / 72.0;
        let width = (page.width().value * scale) as u32;
        let height = (page.height().value * scale) as u32;

        // Create render config
        let render_config = PdfRenderConfig::new()
            .set_target_width(width as i32)
            .set_target_height(height as i32)
            .rotate_if_landscape(PdfPageRenderRotation::None, false);

        // Render to bitmap
        let bitmap = page.render_with_config(&render_config).map_err(|e| {
            ConversionError::PdfRenderError(format!(
                "Failed to render page {}: {}",
                page_idx + 1,
                e
            ))
        })?;

        // Convert to image buffer
        let image_result = bitmap.as_image();

        // Get the raw RGBA bytes
        let rgba_image: RgbaImage = image_result.into_rgba8();

        // Apply background color if not using alpha
        let final_image = if !self.config.use_alpha {
            self.apply_background(rgba_image)
        } else {
            rgba_image
        };

        // Encode to PNG
        let png_data = self.encode_png(&final_image)?;

        Ok(PngPage {
            page_number: page_idx + 1, // 1-indexed
            data: png_data,
            width: final_image.width(),
            height: final_image.height(),
            output_path: None,
        })
    }

    /// Apply background color to transparent areas.
    fn apply_background(&self, mut image: RgbaImage) -> RgbaImage {
        let (r, g, b) = self.config.background_color;

        for pixel in image.pixels_mut() {
            let alpha = pixel[3] as f32 / 255.0;
            if alpha < 1.0 {
                // Blend with background
                let inv_alpha = 1.0 - alpha;
                pixel[0] = ((pixel[0] as f32 * alpha) + (r as f32 * inv_alpha)) as u8;
                pixel[1] = ((pixel[1] as f32 * alpha) + (g as f32 * inv_alpha)) as u8;
                pixel[2] = ((pixel[2] as f32 * alpha) + (b as f32 * inv_alpha)) as u8;
                pixel[3] = 255; // Fully opaque
            }
        }

        image
    }

    /// Encode an image to PNG bytes.
    fn encode_png(&self, image: &RgbaImage) -> Result<Vec<u8>> {
        encode_png_standalone(image)
    }
}

/// Standalone PNG encoding function (Send + Sync safe for parallel execution).
fn encode_png_standalone(image: &RgbaImage) -> Result<Vec<u8>> {
    let mut buffer = Cursor::new(Vec::new());

    // Use the png crate directly for better control over encoding
    let encoder = png::Encoder::new(&mut buffer, image.width(), image.height());
    let mut encoder = encoder;
    encoder.set_color(png::ColorType::Rgba);
    encoder.set_depth(png::BitDepth::Eight);
    encoder.set_compression(png::Compression::Fast); // Use fast compression for throughput

    let mut writer = encoder.write_header().map_err(|e| {
        ConversionError::PngEncodingError(format!("Failed to write PNG header: {}", e))
    })?;

    writer.write_image_data(image.as_raw()).map_err(|e| {
        ConversionError::PngEncodingError(format!("Failed to write PNG data: {}", e))
    })?;

    drop(writer);

    Ok(buffer.into_inner())
}

impl PdfRenderer {
    /// Render pages and save to disk.
    pub fn render_and_save(
        &self,
        pdf_path: &Path,
        output_dir: &Path,
        prefix: &str,
    ) -> Result<Vec<PngPage>> {
        // Ensure output directory exists
        std::fs::create_dir_all(output_dir).map_err(|e| ConversionError::OutputDirError {
            path: output_dir.to_path_buf(),
            message: e.to_string(),
        })?;

        // Render all pages
        let mut pages = self.render_all_pages(pdf_path)?;

        // Save each page
        for page in &mut pages {
            let filename = format!("{}_page_{:04}.png", prefix, page.page_number);
            let output_path = output_dir.join(&filename);

            std::fs::write(&output_path, &page.data).map_err(|e| {
                ConversionError::OutputDirError {
                    path: output_path.clone(),
                    message: e.to_string(),
                }
            })?;

            page.output_path = Some(output_path);
        }

        Ok(pages)
    }

    /// Get an iterator that yields pages one at a time (for streaming).
    pub fn render_pages_iter<'a>(&'a self, pdf_path: &'a Path) -> Result<PageIterator<'a>> {
        let document = self
            .pdfium
            .load_pdf_from_file(pdf_path, None)
            .map_err(|e| ConversionError::PdfRenderError(format!("Failed to load PDF: {}", e)))?;

        let page_count = document.pages().len() as usize;

        Ok(PageIterator {
            renderer: self,
            document,
            current_page: 0,
            total_pages: page_count,
        })
    }
}

/// Iterator over PDF pages for streaming rendering.
pub struct PageIterator<'a> {
    renderer: &'a PdfRenderer,
    document: PdfDocument<'a>,
    current_page: usize,
    total_pages: usize,
}

impl<'a> Iterator for PageIterator<'a> {
    type Item = Result<PngPage>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.current_page >= self.total_pages {
            return None;
        }

        let result = self
            .renderer
            .render_single_page(&self.document, self.current_page);
        self.current_page += 1;
        Some(result)
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        let remaining = self.total_pages - self.current_page;
        (remaining, Some(remaining))
    }
}

impl<'a> ExactSizeIterator for PageIterator<'a> {}

/// Get information about a PDF without rendering it.
pub fn get_pdf_info(pdfium: &Pdfium, pdf_path: &Path) -> Result<PdfInfo> {
    let document = pdfium
        .load_pdf_from_file(pdf_path, None)
        .map_err(|e| ConversionError::PdfRenderError(format!("Failed to load PDF: {}", e)))?;

    let page_count = document.pages().len() as usize;
    let mut pages = Vec::with_capacity(page_count);

    for i in 0..page_count {
        if let Ok(page) = document.pages().get(i as u16) {
            pages.push(PageInfo {
                page_number: i + 1,
                width_points: page.width().value,
                height_points: page.height().value,
            });
        }
    }

    Ok(PdfInfo { page_count, pages })
}

/// Information about a PDF document.
#[derive(Debug, Clone)]
pub struct PdfInfo {
    /// Number of pages.
    pub page_count: usize,
    /// Per-page information.
    pub pages: Vec<PageInfo>,
}

/// Information about a single PDF page.
#[derive(Debug, Clone)]
pub struct PageInfo {
    /// Page number (1-indexed).
    pub page_number: usize,
    /// Width in PDF points (1/72 inch).
    pub width_points: f32,
    /// Height in PDF points (1/72 inch).
    pub height_points: f32,
}

impl PageInfo {
    /// Get width in pixels at a given DPI.
    pub fn width_pixels(&self, dpi: u32) -> u32 {
        ((self.width_points * dpi as f32) / 72.0) as u32
    }

    /// Get height in pixels at a given DPI.
    pub fn height_pixels(&self, dpi: u32) -> u32 {
        ((self.height_points * dpi as f32) / 72.0) as u32
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    // ========== PageInfo tests ==========

    #[test]
    fn test_page_info_dimensions() {
        let info = PageInfo {
            page_number: 1,
            width_points: 612.0,  // US Letter width
            height_points: 792.0, // US Letter height
        };

        // At 72 DPI (1:1)
        assert_eq!(info.width_pixels(72), 612);
        assert_eq!(info.height_pixels(72), 792);

        // At 300 DPI
        assert_eq!(info.width_pixels(300), 2550);
        assert_eq!(info.height_pixels(300), 3300);
    }

    #[test]
    fn test_page_info_a4_dimensions() {
        // A4 paper: 210mm x 297mm = 595.28 x 841.89 points
        let info = PageInfo {
            page_number: 1,
            width_points: 595.28,
            height_points: 841.89,
        };

        // At 72 DPI (1:1)
        assert_eq!(info.width_pixels(72), 595);
        assert_eq!(info.height_pixels(72), 841);

        // At 150 DPI (allow for rounding differences)
        let w150 = info.width_pixels(150);
        let h150 = info.height_pixels(150);
        assert!(w150 == 1240 || w150 == 1239, "Expected ~1240, got {}", w150);
        assert!(
            h150 == 1753 || h150 == 1754,
            "Expected ~1753-1754, got {}",
            h150
        );
    }

    #[test]
    fn test_page_info_clone() {
        let info = PageInfo {
            page_number: 5,
            width_points: 612.0,
            height_points: 792.0,
        };

        let cloned = info.clone();
        assert_eq!(cloned.page_number, 5);
        assert_eq!(cloned.width_points, 612.0);
        assert_eq!(cloned.height_points, 792.0);
    }

    // ========== PdfInfo tests ==========

    #[test]
    fn test_pdf_info_struct() {
        let info = PdfInfo {
            page_count: 3,
            pages: vec![
                PageInfo {
                    page_number: 1,
                    width_points: 612.0,
                    height_points: 792.0,
                },
                PageInfo {
                    page_number: 2,
                    width_points: 612.0,
                    height_points: 792.0,
                },
                PageInfo {
                    page_number: 3,
                    width_points: 792.0, // Landscape
                    height_points: 612.0,
                },
            ],
        };

        assert_eq!(info.page_count, 3);
        assert_eq!(info.pages.len(), 3);
        // Check landscape page
        assert!(info.pages[2].width_points > info.pages[2].height_points);
    }

    #[test]
    fn test_pdf_info_empty() {
        let info = PdfInfo {
            page_count: 0,
            pages: vec![],
        };

        assert_eq!(info.page_count, 0);
        assert!(info.pages.is_empty());
    }

    // ========== encode_png_standalone tests ==========

    #[test]
    fn test_encode_png_small_image() {
        // Create a 10x10 red image
        let mut image = RgbaImage::new(10, 10);
        for pixel in image.pixels_mut() {
            *pixel = Rgba([255, 0, 0, 255]); // Red
        }

        let png_data = encode_png_standalone(&image).unwrap();

        // Verify it's a valid PNG (check magic bytes)
        assert!(png_data.len() > 8);
        assert_eq!(
            &png_data[0..8],
            &[0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]
        );
    }

    #[test]
    fn test_encode_png_with_transparency() {
        // Create a 5x5 semi-transparent image
        let mut image = RgbaImage::new(5, 5);
        for pixel in image.pixels_mut() {
            *pixel = Rgba([0, 255, 0, 128]); // Semi-transparent green
        }

        let png_data = encode_png_standalone(&image).unwrap();

        // Verify PNG header
        assert_eq!(
            &png_data[0..8],
            &[0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]
        );
    }

    #[test]
    fn test_encode_png_1x1_pixel() {
        // Minimal image
        let mut image = RgbaImage::new(1, 1);
        image.put_pixel(0, 0, Rgba([100, 150, 200, 255]));

        let png_data = encode_png_standalone(&image).unwrap();
        assert!(!png_data.is_empty());
        assert_eq!(
            &png_data[0..8],
            &[0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]
        );
    }

    #[test]
    fn test_encode_png_gradient() {
        // Create a gradient image
        let mut image = RgbaImage::new(256, 1);
        for x in 0..256 {
            image.put_pixel(x, 0, Rgba([x as u8, x as u8, x as u8, 255]));
        }

        let png_data = encode_png_standalone(&image).unwrap();
        assert!(!png_data.is_empty());
    }

    // ========== RenderConfig validation tests ==========

    #[test]
    fn test_render_config_validation_valid() {
        let config = RenderConfig::with_dpi(150);
        assert!(config.validate().is_ok());
    }

    #[test]
    fn test_render_config_validation_zero_dpi() {
        let mut config = RenderConfig::default();
        config.dpi = 0;

        let result = config.validate();
        assert!(result.is_err());
    }

    #[test]
    fn test_render_config_validation_excessive_dpi() {
        let mut config = RenderConfig::default();
        config.dpi = 5000; // Way too high

        let result = config.validate();
        assert!(result.is_err());
    }

    #[test]
    fn test_render_config_validation_zero_threads() {
        let mut config = RenderConfig::default();
        config.render_threads = 0;

        let result = config.validate();
        assert!(result.is_err());
    }

    // ========== PdfRenderer tests (require pdfium) ==========

    #[test]
    fn test_renderer_creation_with_invalid_config() {
        let mut config = RenderConfig::default();
        config.dpi = 0;

        let result = PdfRenderer::new(config);
        assert!(result.is_err());
    }

    #[test]
    fn test_renderer_dpi_accessor() {
        // This will only work if pdfium is installed
        let config = RenderConfig::with_dpi(200);
        match PdfRenderer::new(config) {
            Ok(renderer) => {
                assert_eq!(renderer.dpi(), 200);
            }
            Err(ConversionError::PdfiumError(_)) => {
                // pdfium not installed, skip test
            }
            Err(e) => panic!("Unexpected error: {:?}", e),
        }
    }
}
