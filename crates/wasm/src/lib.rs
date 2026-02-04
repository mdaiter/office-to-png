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
//!
//! # Example (JavaScript)
//!
//! ```javascript
//! import init, { DocumentRenderer } from 'office-to-png-wasm';
//!
//! async function renderDocument() {
//!     await init();
//!     
//!     const renderer = new DocumentRenderer('canvas-id');
//!     
//!     // Load a file from user input
//!     const file = document.getElementById('file-input').files[0];
//!     const data = new Uint8Array(await file.arrayBuffer());
//!     
//!     if (file.name.endsWith('.docx')) {
//!         const info = renderer.load_docx(data);
//!         console.log('Loaded document with', info.page_count, 'pages');
//!         renderer.render_page(0); // Render first page
//!     } else if (file.name.endsWith('.xlsx')) {
//!         const info = renderer.load_xlsx(data);
//!         console.log('Loaded spreadsheet with', info.page_count, 'sheets');
//!         renderer.render_sheet(0); // Render first sheet
//!     }
//! }
//! ```

use wasm_bindgen::prelude::*;
use web_sys::console;

pub mod docx_renderer;
pub mod fonts;
pub mod renderer;
pub mod styles;
pub mod text_layout;
pub mod text_shaper;
pub mod xlsx_grid_renderer;
pub mod xlsx_renderer;

// Re-export renderer types
pub use renderer::{Canvas2DRenderer, Color, RenderBackend};
pub use xlsx_grid_renderer::XlsxGridRenderer;

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
#[wasm_bindgen]
pub struct DocumentRenderer {
    canvas_id: String,
    docx_doc: Option<docx_renderer::DocxDocument>,
    xlsx_doc: Option<xlsx_renderer::XlsxDocument>,
    xlsx_sheet_data: Option<xlsx_renderer::SheetData>,
    current_page: usize,
    current_sheet: usize,
    xlsx_renderer: xlsx_grid_renderer::XlsxGridRenderer,
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
