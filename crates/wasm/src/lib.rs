//! Canvas 2D-based Office document rendering for browsers.
//!
//! This module provides client-side rendering of Office documents
//! using the HTML5 Canvas 2D API for maximum browser compatibility.
//!
//! # Features
//!
//! - **DOCX rendering**: Paragraphs, text formatting, tables
//! - **XLSX rendering**: Cells, styling, borders, merged cells
//! - **Canvas 2D backend**: Works in all modern browsers
//! - **PNG export**: Export rendered pages as PNG images
//! - **Web Worker support**: Heavy parsing runs off main thread
//!
//! # Example (JavaScript) - Synchronous Mode
//!
//! ```javascript
//! import init, { DocumentRenderer } from 'office-to-png-wasm';
//!
//! async function renderDocument() {
//!     await init();
//!     
//!     const renderer = new DocumentRenderer('canvas-id');
//!     
//!     const file = document.getElementById('file-input').files[0];
//!     const data = new Uint8Array(await file.arrayBuffer());
//!     
//!     if (file.name.endsWith('.docx')) {
//!         const info = renderer.load_docx(data);
//!         renderer.render_page(0);
//!     } else if (file.name.endsWith('.xlsx')) {
//!         const info = renderer.load_xlsx(data);
//!         renderer.render_sheet(0);
//!     }
//! }
//! ```
//!
//! # Example (JavaScript) - Web Worker Mode
//!
//! ```javascript
//! // Main thread
//! import init, { DocumentRenderer } from 'office-to-png-wasm';
//!
//! const worker = new Worker('./worker.js');
//! const renderer = new DocumentRenderer('canvas-id');
//!
//! worker.onmessage = (e) => {
//!     const { buffers, sheetIndex } = e.data;
//!     const [header, data, ...images] = buffers;
//!     renderer.render_sheet_from_bytes(new Uint8Array(data), images, sheetIndex);
//! };
//!
//! // Request sheet parse from worker
//! worker.postMessage({ type: 'parse_sheet', docBytes, sheetIndex });
//! ```

use js_sys::Array;
use std::collections::HashMap;
use wasm_bindgen::prelude::*;
use web_sys::console;

pub mod docx_renderer;
pub mod fonts;
pub mod render_data;
pub mod renderer;
pub mod styles;
pub mod text_layout;
pub mod text_shaper;
pub mod worker_api;
pub mod xlsx_grid_renderer;
pub mod xlsx_renderer;

// Re-export renderer types
pub use render_data::{DocxPageData, RenderPrimitive, XlsxSheetRenderData};
pub use renderer::{Canvas2DRenderer, Color, RenderBackend};
pub use xlsx_grid_renderer::XlsxGridRenderer;

// Re-export worker API functions and types
pub use worker_api::{
    worker_get_docx_page_count, worker_get_document_info, worker_get_xlsx_sheet_names,
    worker_parse_docx_page, worker_parse_xlsx_sheet, worker_parse_xlsx_sheet_with_size,
    WorkerDocumentHolder,
};

/// Initialize the WASM module.
#[wasm_bindgen(start)]
pub fn init() {
    // Set up better panic messages
    console_error_panic_hook::set_once();
    console::log_1(&"office-to-png-wasm initialized".into());
}

/// Supported document types.
#[wasm_bindgen]
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum DocumentType {
    Docx,
    Xlsx,
    Unknown,
}

/// Information about a loaded document.
#[wasm_bindgen]
pub struct DocumentInfo {
    doc_type: DocumentType,
    page_count: usize,
    title: Option<String>,
}

#[wasm_bindgen]
impl DocumentInfo {
    #[wasm_bindgen(getter)]
    pub fn doc_type(&self) -> DocumentType {
        self.doc_type
    }
    
    #[wasm_bindgen(getter)]
    pub fn page_count(&self) -> usize {
        self.page_count
    }
    
    #[wasm_bindgen(getter)]
    pub fn title(&self) -> Option<String> {
        self.title.clone()
    }
}

/// Document renderer with full Canvas 2D rendering support.
///
/// Supports both synchronous mode (document loaded locally) and
/// async mode (render data received from Web Worker).
#[wasm_bindgen]
pub struct DocumentRenderer {
    canvas_id: String,
    docx_doc: Option<docx_renderer::DocxDocument>,
    xlsx_doc: Option<xlsx_renderer::XlsxDocument>,
    xlsx_sheet_data: Option<xlsx_renderer::SheetData>,
    current_page: usize,
    current_sheet: usize,
    xlsx_renderer: xlsx_grid_renderer::XlsxGridRenderer,

    // Render versioning for staleness detection
    render_version: u64,

    // Caching for pre-computed render data (bincode bytes)
    page_cache: HashMap<usize, Vec<u8>>,
    sheet_cache: HashMap<usize, Vec<u8>>,
    page_image_cache: HashMap<usize, Vec<Vec<u8>>>,
    sheet_image_cache: HashMap<usize, Vec<Vec<u8>>>,
    cache_enabled: bool,
}

#[wasm_bindgen]
impl DocumentRenderer {
    #[wasm_bindgen(constructor)]
    pub fn new(canvas_id: &str) -> Self {
        console::log_1(&format!("Creating DocumentRenderer for canvas: {}", canvas_id).into());
        Self {
            canvas_id: canvas_id.to_string(),
            docx_doc: None,
            xlsx_doc: None,
            xlsx_sheet_data: None,
            current_page: 0,
            current_sheet: 0,
            xlsx_renderer: xlsx_grid_renderer::XlsxGridRenderer::new(),
            render_version: 0,
            page_cache: HashMap::new(),
            sheet_cache: HashMap::new(),
            page_image_cache: HashMap::new(),
            sheet_image_cache: HashMap::new(),
            cache_enabled: true,
        }
    }
    
    /// Load a Word document (.docx) from bytes.
    pub fn load_docx(&mut self, data: &[u8]) -> Result<DocumentInfo, JsValue> {
        let doc = docx_renderer::DocxDocument::from_bytes(data)
            .map_err(|e| JsValue::from_str(&format!("Failed to load DOCX: {}", e)))?;
        
        let info = DocumentInfo {
            doc_type: DocumentType::Docx,
            page_count: doc.page_count(),
            title: doc.title().map(String::from),
        };
        
        self.docx_doc = Some(doc);
        self.xlsx_doc = None;
        self.xlsx_sheet_data = None;
        self.current_page = 0;
        self.current_sheet = 0;
        
        Ok(info)
    }
    
    /// Load an Excel spreadsheet (.xlsx) from bytes.
    pub fn load_xlsx(&mut self, data: &[u8]) -> Result<DocumentInfo, JsValue> {
        let doc = xlsx_renderer::XlsxDocument::from_bytes(data)
            .map_err(|e| JsValue::from_str(&format!("Failed to load XLSX: {}", e)))?;
        
        let info = DocumentInfo {
            doc_type: DocumentType::Xlsx,
            page_count: doc.sheet_count(),
            title: doc.title().map(String::from),
        };
        
        // Pre-load first sheet data
        let sheet_data = doc.get_styled_sheet_data(0);
        
        self.xlsx_doc = Some(doc);
        self.xlsx_sheet_data = sheet_data;
        self.docx_doc = None;
        self.current_page = 0;
        self.current_sheet = 0;
        
        Ok(info)
    }
    
    /// Render the current page (for DOCX).
    pub fn render_page(&self, page: usize) -> Result<(), JsValue> {
        let renderer = Canvas2DRenderer::new(&self.canvas_id)
            .map_err(|e| JsValue::from_str(&e))?;
        
        if let Some(ref doc) = self.docx_doc {
            doc.render(&renderer, page)
                .map_err(|e| JsValue::from_str(&format!("Render error: {}", e)))?;
        }
        
        Ok(())
    }
    
    /// Render a sheet (for XLSX).
    pub fn render_sheet(&mut self, sheet_index: usize) -> Result<(), JsValue> {
        let renderer = Canvas2DRenderer::new(&self.canvas_id)
            .map_err(|e| JsValue::from_str(&e))?;
        
        if let Some(ref doc) = self.xlsx_doc {
            // Load sheet data if needed
            if self.current_sheet != sheet_index || self.xlsx_sheet_data.is_none() {
                self.xlsx_sheet_data = doc.get_styled_sheet_data(sheet_index);
                self.current_sheet = sheet_index;
            }
            
            if let Some(ref sheet_data) = self.xlsx_sheet_data {
                self.xlsx_renderer.render(&renderer, sheet_data)
                    .map_err(|e| JsValue::from_str(&e))?;
            }
        }
        
        Ok(())
    }
    
    /// Set zoom level (0.25 to 2.0).
    pub fn set_zoom(&mut self, zoom: f32) {
        self.xlsx_renderer.set_zoom(zoom);
    }
    
    /// Set scroll position for XLSX.
    pub fn set_scroll(&mut self, x: f32, y: f32) {
        self.xlsx_renderer.set_scroll(x, y);
    }
    
    /// Export current view as PNG bytes.
    pub fn export_png(&self) -> Result<Vec<u8>, JsValue> {
        let renderer = Canvas2DRenderer::new(&self.canvas_id)
            .map_err(|e| JsValue::from_str(&e))?;
        
        renderer.export_png()
            .map_err(|e| JsValue::from_str(&e))
    }
    
    /// Get the current page number.
    #[wasm_bindgen(getter)]
    pub fn current_page(&self) -> usize {
        self.current_page
    }
    
    /// Get the current sheet index.
    #[wasm_bindgen(getter)]
    pub fn current_sheet(&self) -> usize {
        self.current_sheet
    }
    
    /// Get the total number of pages.
    #[wasm_bindgen(getter)]
    pub fn page_count(&self) -> usize {
        if let Some(ref doc) = self.docx_doc {
            doc.page_count()
        } else if let Some(ref doc) = self.xlsx_doc {
            doc.sheet_count()
        } else {
            0
        }
    }
    
    /// Get sheet names (for Excel documents).
    pub fn get_sheet_names(&self) -> Vec<String> {
        if let Some(ref doc) = self.xlsx_doc {
            doc.sheet_names().iter().map(|s| s.to_string()).collect()
        } else {
            vec![]
        }
    }
    
    /// Get cell data for a sheet (for Excel documents).
    pub fn get_sheet_data(&self, sheet_index: usize) -> Result<JsValue, JsValue> {
        if let Some(ref doc) = self.xlsx_doc {
            if let Some(data) = doc.get_cell_data(sheet_index) {
                serde_wasm_bindgen::to_value(&data)
                    .map_err(|e| JsValue::from_str(&format!("Serialization error: {}", e)))
            } else {
                Err(JsValue::from_str("Sheet not found"))
            }
        } else {
            Err(JsValue::from_str("No Excel document loaded"))
        }
    }
    
    /// Get styled sheet data (includes formatting).
    pub fn get_styled_sheet_data(&self, sheet_index: usize) -> Result<JsValue, JsValue> {
        if let Some(ref doc) = self.xlsx_doc {
            if let Some(data) = doc.get_styled_sheet_data(sheet_index) {
                serde_wasm_bindgen::to_value(&data)
                    .map_err(|e| JsValue::from_str(&format!("Serialization error: {}", e)))
            } else {
                Err(JsValue::from_str("Sheet not found"))
            }
        } else {
            Err(JsValue::from_str("No Excel document loaded"))
        }
    }

    // ========================================================================
    // Web Worker API: Render Versioning and Caching
    // ========================================================================

    /// Start a new render request, returns version ID for staleness detection.
    /// Call this before sending a parse request to the worker.
    pub fn begin_render(&mut self) -> u64 {
        self.render_version += 1;
        self.render_version
    }

    /// Check if a render version is still current.
    /// If false, the render result should be discarded.
    pub fn is_current_version(&self, version: u64) -> bool {
        self.render_version == version
    }

    /// Get the current render version.
    #[wasm_bindgen(getter)]
    pub fn render_version(&self) -> u64 {
        self.render_version
    }

    /// Check if a page is cached.
    pub fn has_cached_page(&self, page_index: usize) -> bool {
        self.page_cache.contains_key(&page_index)
    }

    /// Check if a sheet is cached.
    pub fn has_cached_sheet(&self, sheet_index: usize) -> bool {
        self.sheet_cache.contains_key(&sheet_index)
    }

    /// Get cached page data bytes (if available).
    pub fn get_cached_page_bytes(&self, page_index: usize) -> Option<Vec<u8>> {
        self.page_cache.get(&page_index).cloned()
    }

    /// Get cached sheet data bytes (if available).
    pub fn get_cached_sheet_bytes(&self, sheet_index: usize) -> Option<Vec<u8>> {
        self.sheet_cache.get(&sheet_index).cloned()
    }

    /// Get cached page image bytes as JS Array.
    pub fn get_cached_page_images(&self, page_index: usize) -> Array {
        let result = Array::new();
        if let Some(images) = self.page_image_cache.get(&page_index) {
            for img in images {
                let arr = js_sys::Uint8Array::from(&img[..]);
                result.push(&arr.buffer());
            }
        }
        result
    }

    /// Get cached sheet image bytes as JS Array.
    pub fn get_cached_sheet_images(&self, sheet_index: usize) -> Array {
        let result = Array::new();
        if let Some(images) = self.sheet_image_cache.get(&sheet_index) {
            for img in images {
                let arr = js_sys::Uint8Array::from(&img[..]);
                result.push(&arr.buffer());
            }
        }
        result
    }

    /// Enable or disable caching.
    pub fn set_cache_enabled(&mut self, enabled: bool) {
        self.cache_enabled = enabled;
        if !enabled {
            self.clear_cache();
        }
    }

    /// Check if caching is enabled.
    #[wasm_bindgen(getter)]
    pub fn cache_enabled(&self) -> bool {
        self.cache_enabled
    }

    /// Clear all cached render data.
    pub fn clear_cache(&mut self) {
        self.page_cache.clear();
        self.sheet_cache.clear();
        self.page_image_cache.clear();
        self.sheet_image_cache.clear();
    }

    // ========================================================================
    // Web Worker API: Render from Pre-computed Data
    // ========================================================================

    /// Render a DOCX page from pre-computed binary data.
    ///
    /// # Arguments
    /// * `data_bytes` - bincode-serialized DocxPageData
    /// * `image_buffers` - Array of ArrayBuffers containing image data
    /// * `page_index` - The page index being rendered
    pub fn render_page_from_bytes(
        &mut self,
        data_bytes: &[u8],
        image_buffers: Array,
        page_index: usize,
    ) -> Result<(), JsValue> {
        use crate::render_data::DocxPageData;

        // Cache if enabled
        if self.cache_enabled {
            self.page_cache.insert(page_index, data_bytes.to_vec());
            let mut images = Vec::new();
            for i in 0..image_buffers.length() {
                let buf = image_buffers.get(i);
                let arr = js_sys::Uint8Array::new(&buf);
                images.push(arr.to_vec());
            }
            self.page_image_cache.insert(page_index, images);
        }

        // Deserialize
        let page_data: DocxPageData = bincode::deserialize(data_bytes)
            .map_err(|e| JsValue::from_str(&format!("Deserialize error: {}", e)))?;

        // Collect image bytes from ArrayBuffers
        let mut image_bytes = Vec::new();
        for i in 0..image_buffers.length() {
            let buf = image_buffers.get(i);
            let arr = js_sys::Uint8Array::new(&buf);
            image_bytes.push(arr.to_vec());
        }

        // Create canvas renderer and draw primitives
        let canvas = Canvas2DRenderer::new(&self.canvas_id)
            .map_err(|e| JsValue::from_str(&e))?;

        self.draw_primitives(&canvas, &page_data.primitives, &image_bytes)?;

        self.current_page = page_index;
        Ok(())
    }

    /// Render an XLSX sheet from pre-computed binary data.
    ///
    /// # Arguments
    /// * `data_bytes` - bincode-serialized XlsxSheetRenderData
    /// * `image_buffers` - Array of ArrayBuffers containing image data
    /// * `sheet_index` - The sheet index being rendered
    pub fn render_sheet_from_bytes(
        &mut self,
        data_bytes: &[u8],
        image_buffers: Array,
        sheet_index: usize,
    ) -> Result<(), JsValue> {
        use crate::render_data::XlsxSheetRenderData;

        // Cache if enabled
        if self.cache_enabled {
            self.sheet_cache.insert(sheet_index, data_bytes.to_vec());
            let mut images = Vec::new();
            for i in 0..image_buffers.length() {
                let buf = image_buffers.get(i);
                let arr = js_sys::Uint8Array::new(&buf);
                images.push(arr.to_vec());
            }
            self.sheet_image_cache.insert(sheet_index, images);
        }

        // Deserialize
        let sheet_data: XlsxSheetRenderData = bincode::deserialize(data_bytes)
            .map_err(|e| JsValue::from_str(&format!("Deserialize error: {}", e)))?;

        // Collect image bytes
        let mut image_bytes = Vec::new();
        for i in 0..image_buffers.length() {
            let buf = image_buffers.get(i);
            let arr = js_sys::Uint8Array::new(&buf);
            image_bytes.push(arr.to_vec());
        }

        // Create canvas renderer and draw primitives
        let canvas = Canvas2DRenderer::new(&self.canvas_id)
            .map_err(|e| JsValue::from_str(&e))?;

        self.draw_primitives(&canvas, &sheet_data.primitives, &image_bytes)?;

        self.current_sheet = sheet_index;
        Ok(())
    }

    /// Draw pre-computed render primitives to the canvas.
    fn draw_primitives(
        &self,
        canvas: &Canvas2DRenderer,
        primitives: &[render_data::RenderPrimitive],
        image_bytes: &[Vec<u8>],
    ) -> Result<(), JsValue> {
        use crate::render_data::RenderPrimitive;
        use crate::text_layout::TextStyle;

        for prim in primitives {
            match prim {
                RenderPrimitive::Clear(color) => {
                    canvas
                        .clear(Color::from_rgba_array(*color))
                        .map_err(|e| JsValue::from_str(&e))?;
                }
                RenderPrimitive::Text(t) => {
                    let mut style = TextStyle::new(&t.font_family, t.font_size);
                    style.bold = t.bold;
                    style.italic = t.italic;
                    style.underline = t.underline;
                    style.strikethrough = t.strikethrough;
                    style.color = t.color;
                    style.background = t.background;

                    canvas
                        .draw_text(&t.text, t.x, t.y, &style)
                        .map_err(|e| JsValue::from_str(&e))?;
                }
                RenderPrimitive::Line(l) => {
                    canvas
                        .draw_line(l.x1, l.y1, l.x2, l.y2, Color::from_rgba_array(l.color), l.width)
                        .map_err(|e| JsValue::from_str(&e))?;
                }
                RenderPrimitive::Rect(r) => {
                    if let Some(fill) = r.fill {
                        canvas
                            .fill_rect(r.x, r.y, r.width, r.height, Color::from_rgba_array(fill))
                            .map_err(|e| JsValue::from_str(&e))?;
                    }
                    if let Some((color, width)) = r.stroke {
                        let border = crate::renderer::BorderStyle {
                            color: Color::from_rgba_array(color),
                            width,
                            ..Default::default()
                        };
                        canvas
                            .stroke_rect(r.x, r.y, r.width, r.height, &border)
                            .map_err(|e| JsValue::from_str(&e))?;
                    }
                }
                RenderPrimitive::Image(img) => {
                    if let Some(bytes) = image_bytes.get(img.image_index) {
                        // Decode image
                        let decoded = image::load_from_memory(bytes)
                            .map_err(|e| JsValue::from_str(&format!("Image decode error: {}", e)))?;
                        let rgba = decoded.to_rgba8();
                        let (w, h) = rgba.dimensions();

                        canvas
                            .draw_image(
                                &rgba.into_raw(),
                                w,
                                h,
                                img.x,
                                img.y,
                                img.dest_width,
                                img.dest_height,
                            )
                            .map_err(|e| JsValue::from_str(&e))?;
                    }
                }
                RenderPrimitive::Clip { x, y, width, height } => {
                    canvas
                        .clip(*x, *y, *width, *height)
                        .map_err(|e| JsValue::from_str(&e))?;
                }
                RenderPrimitive::Save => {
                    canvas.save().map_err(|e| JsValue::from_str(&e))?;
                }
                RenderPrimitive::Restore => {
                    canvas.restore().map_err(|e| JsValue::from_str(&e))?;
                }
                RenderPrimitive::ResetClip => {
                    // Reset is typically done via restore
                }
            }
        }

        Ok(())
    }
}

/// Check if WebGPU is supported (always returns status for now).
#[wasm_bindgen]
pub async fn is_webgpu_supported() -> bool {
    // Check if the browser supports WebGPU
    let window = match web_sys::window() {
        Some(w) => w,
        None => return false,
    };
    
    let navigator = window.navigator();
    
    // Check if gpu property exists on navigator
    js_sys::Reflect::has(&navigator, &"gpu".into()).unwrap_or(false)
}

/// Get supported file extensions.
#[wasm_bindgen]
pub fn supported_extensions() -> Vec<String> {
    vec![
        "docx".to_string(),
        "xlsx".to_string(),
    ]
}
