# office-to-png

A high-performance Rust library for converting Microsoft Office documents (.docx, .xlsx) to PNG images.

## Features

- **Dual rendering paths**:
  - **Server-side (core crate)**: LibreOffice -> PDF -> pdfium -> PNG
  - **Browser-side (wasm crate)**: Direct parsing -> Canvas 2D/WebGPU -> PNG

- **Full document support**:
  - DOCX: Text, paragraphs, tables, images, styling (bold, italic, colors, etc.)
  - XLSX: Cells, grids, formatting, merged cells, multiple sheets

- **High performance**:
  - Parallel page rendering with rayon
  - LibreOffice instance pooling for batch conversions
  - WebGPU acceleration in browsers (optional)
  - **Web Worker support** for non-blocking UI during document parsing
  - Intelligent caching of parsed documents and rendered sheets

- **Cross-platform**: Works on macOS, Linux, and Windows

## Crates

| Crate | Description |
|-------|-------------|
| `office-to-png-core` | Server-side conversion using LibreOffice + pdfium |
| `office-to-png-server` | HTTP API server for document conversion (REST API) |
| `office-to-png-wasm` | Browser-side rendering with Canvas 2D or WebGPU |
| `office-to-png-python` | Python bindings via PyO3 |

## Installation

### Core Library (Server-side)

Add to your `Cargo.toml`:

```toml
[dependencies]
office-to-png-core = { git = "https://github.com/mdaiter/office-to-png" }
```

**Requirements**:
- LibreOffice installed (`soffice` in PATH or specify path)
- pdfium library (download from [pdfium-binaries](https://github.com/nicovank/pdfium-cmake))

### WASM Library (Browser-side)

```toml
[dependencies]
office-to-png-wasm = { git = "https://github.com/mdaiter/office-to-png" }
```

To enable the optional WebGPU backend:
```toml
[dependencies]
office-to-png-wasm = { git = "https://github.com/mdaiter/office-to-png", features = ["webgpu"] }
```

Build with:
```bash
wasm-pack build crates/wasm --target web
```

### Python

**Note**: Requires Python 3.8-3.13 (Python 3.14+ not yet supported by PyO3).

```bash
cd crates/python

# Create and activate a virtual environment
python3 -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate

# Install maturin and build
pip install maturin
maturin develop

# Verify installation
python -c "from office_to_png import is_libreoffice_available; print('LibreOffice:', is_libreoffice_available())"
```

## Usage

### Rust (Server-side)

```rust
use office_to_png_core::{Converter, ConversionRequest};
use std::path::PathBuf;

#[tokio::main]
async fn main() -> anyhow::Result<()> {
    // Create converter with default settings
    let converter = Converter::builder()
        .pool_size(4)
        .dpi(150)
        .build()
        .await?;

    // Convert a single file
    let request = ConversionRequest::new(
        PathBuf::from("document.docx"),
        PathBuf::from("output/"),
    );
    
    let result = converter.convert(request).await?;
    
    for page in result.pages {
        println!("Generated: {:?}", page.path);
    }
    
    Ok(())
}
```

### Rust (Browser/WASM)

```rust
use office_to_png_wasm::{DocxDocument, Canvas2DRenderer};

// Parse document
let doc = DocxDocument::from_bytes(&docx_bytes)?;

// Create renderer targeting a canvas element
let renderer = Canvas2DRenderer::new("canvas-id")?;

// Render page 0
doc.render(&renderer, 0)?;

// Export to PNG
let png_bytes = renderer.export_png()?;
```

### JavaScript (via WASM)

```javascript
import init, { DocumentRenderer } from './pkg/office_to_png_wasm.js';

await init();

// Create renderer targeting a canvas element
const renderer = new DocumentRenderer('my-canvas');

// Load and render a DOCX file
const response = await fetch('document.docx');
const bytes = new Uint8Array(await response.arrayBuffer());
const info = renderer.load_docx(bytes);
console.log(`Loaded ${info.page_count} pages`);

// Render page 0
renderer.render_page(0);

// Export as PNG
const pngBytes = renderer.export_png();

// For XLSX files:
// const info = renderer.load_xlsx(bytes);
// renderer.render_sheet(0);
// const sheetNames = renderer.get_sheet_names();
```

### JavaScript with Web Worker (Recommended for Large Documents)

For better performance with large documents, use the Web Worker API to avoid blocking the main thread:

```javascript
import init, { DocumentRenderer, WorkerDocumentHolder } from './pkg/office_to_png_wasm.js';

// In your Web Worker (worker.js):
let docHolder = null;

self.onmessage = async (e) => {
    const { type, data } = e.data;
    
    if (type === 'init') {
        const wasm = await import('./pkg/office_to_png_wasm.js');
        await wasm.default();
        docHolder = new wasm.WorkerDocumentHolder();
        self.postMessage({ type: 'ready' });
    }
    
    if (type === 'load') {
        docHolder.load(data.bytes, data.docType);
        self.postMessage({ type: 'loaded' });
    }
    
    if (type === 'parse_sheet') {
        // Parse sheet in worker (non-blocking)
        const buffers = docHolder.parse_xlsx_sheet(
            data.sheetIndex, 
            data.requestId,
            data.canvasWidth,
            data.canvasHeight
        );
        // Transfer buffers back to main thread (zero-copy)
        self.postMessage({ type: 'parsed', buffers }, buffers);
    }
};

// In your main thread:
const worker = new Worker('./worker.js', { type: 'module' });
const renderer = new DocumentRenderer('my-canvas');

worker.onmessage = (e) => {
    if (e.data.type === 'parsed') {
        const dataBytes = new Uint8Array(e.data.buffers[1]);
        const imageBuffers = e.data.buffers.slice(2);
        renderer.render_sheet_from_bytes(dataBytes, imageBuffers, sheetIndex);
    }
};
```

### Python

```python
import asyncio
from office_to_png import OfficeConverter

async def main():
    # Create converter with 4 LibreOffice instances
    converter = OfficeConverter(pool_size=4, dpi=150)
    
    # Convert a single file
    result = await converter.convert("document.docx", "./output")
    print(f"Converted {result.page_count} pages in {result.duration_secs:.2f}s")
    
    for path in result.output_paths:
        print(f"  Generated: {path}")
    
    # Batch convert with progress callback
    def on_progress(p):
        print(f"[{p.file_index + 1}/{p.total_files}] {p.stage}: {p.current_file}")
    
    batch_result = await converter.convert_batch(
        ["doc1.docx", "doc2.xlsx", "doc3.docx"],
        "./output",
        progress_callback=on_progress
    )
    print(f"Batch: {batch_result.success_count} succeeded, {batch_result.failure_count} failed")
    
    # Cleanup
    await converter.shutdown()

asyncio.run(main())
```

See `examples/python_quickstart.py` for a complete walkthrough.

### HTTP API Server

The server crate provides a REST API for document conversion with **100% fidelity** using LibreOffice:

```bash
# Run the server (uses local filesystem storage by default)
cargo run -p office-to-png-server

# Run on a custom port
PORT=8085 cargo run -p office-to-png-server

# Or with Docker
docker-compose -f deploy/docker-compose.yml up
```

**API Endpoints:**

| Method | Path | Description |
|--------|------|-------------|
| `POST` | `/api/upload` | Upload document, returns job ID |
| `GET` | `/api/job/:id` | Get job status |
| `GET` | `/api/job/:id/pdf` | Download converted PDF |
| `GET` | `/api/job/:id/png/:page` | Download specific PNG page (1-indexed) |
| `DELETE` | `/api/job/:id` | Delete job and files |
| `GET` | `/api/jobs` | List all jobs |
| `GET` | `/health` | Health check |

**Example:**

```bash
# Upload a document
curl -X POST -F "file=@document.docx" http://localhost:8080/api/upload
# Returns: {"job_id": "abc123", "status": "pending", ...}

# Check status (poll until "completed")
curl http://localhost:8080/api/job/abc123
# Returns: {"job_id": "abc123", "status": "completed", "page_count": 5, ...}

# Download PNG page 1
curl http://localhost:8080/api/job/abc123/png/1 -o page1.png
```

**Environment Variables:**

| Variable | Default | Description |
|----------|---------|-------------|
| `PORT` | `8080` | Server port |
| `HOST` | `0.0.0.0` | Server host |
| `STORAGE_BACKEND` | auto | `local` or `s3` (auto-detects based on `S3_BUCKET`) |
| `STORAGE_DIR` | system temp | Directory for local storage |
| `S3_BUCKET` | - | S3 bucket (enables S3 mode when set) |
| `S3_ENDPOINT` | - | Custom S3 endpoint (for LocalStack/MinIO) |
| `AWS_REGION` | - | AWS region |
| `POOL_SIZE` | `2` | LibreOffice instance count |
| `DPI` | `150` | PNG rendering DPI |

**Storage Backends:**

- **Local (default)**: Files stored on local filesystem. No AWS setup required - perfect for development.
- **S3**: Set `S3_BUCKET` to enable. Supports LocalStack/MinIO via `S3_ENDPOINT`.

## Configuration

### Converter Options

| Option | Default | Description |
|--------|---------|-------------|
| `pool_size` | 2 | Number of LibreOffice instances |
| `dpi` | 150 | Output image resolution |
| `render_threads` | CPU cores | Parallel rendering threads |
| `conversion_timeout` | 60s | Timeout per document |
| `soffice_path` | Auto-detect | Path to LibreOffice |

### Render Options

| Option | Default | Description |
|--------|---------|-------------|
| `dpi` | 150 | Dots per inch |
| `background` | White | Page background color |
| `png_compression` | 6 | PNG compression level (0-9) |

### Feature Flags (WASM crate)

| Feature | Description |
|---------|-------------|
| `webgpu` | Enable WebGPU rendering backend for GPU-accelerated rendering. Falls back to Canvas 2D if WebGPU is unavailable in the browser. |

## Architecture

```
office-to-png/
├── crates/
│   ├── core/           # Server-side: LibreOffice + pdfium
│   │   ├── converter.rs    # Main conversion orchestration
│   │   ├── pool.rs         # LibreOffice instance pooling
│   │   └── pdf_renderer.rs # pdfium-based PDF to PNG
│   │
│   ├── server/         # HTTP API server
│   │   ├── handlers.rs     # API endpoint handlers
│   │   ├── job.rs          # Job tracking and management
│   │   └── storage.rs      # Storage backends (local filesystem + S3)
│   │
│   ├── wasm/           # Browser-side: Direct rendering
│   │   ├── docx_renderer.rs  # DOCX parsing + rendering
│   │   ├── xlsx_renderer.rs  # XLSX parsing + rendering
│   │   ├── xlsx_grid_renderer.rs  # XLSX grid/cell rendering
│   │   ├── renderer/
│   │   │   ├── traits.rs     # RenderBackend trait
│   │   │   ├── canvas2d.rs   # Canvas 2D implementation
│   │   │   └── webgpu.rs     # WebGPU implementation
│   │   ├── worker_api.rs     # Web Worker API for async parsing
│   │   ├── render_data.rs    # Serializable render primitives
│   │   └── text_shaper.rs    # cosmic-text integration
│   │
│   └── python/         # Python bindings via PyO3
│
├── deploy/             # Deployment configurations
│   ├── Dockerfile      # Server container
│   └── docker-compose.yml  # Local dev with LocalStack
│
├── demo/               # Interactive browser demo
│   ├── index.html      # Demo UI with drag-and-drop
│   ├── test.html       # Minimal WASM test page
│   ├── test-server.html # Server HD mode test page
│   ├── worker.js       # Web Worker for async document parsing
│   └── pkg/            # Built WASM package
│
├── tests/
│   └── fixtures/       # Test documents
│
└── examples/           # Usage examples
```

### Web Worker Architecture

For optimal browser performance, the WASM crate supports a Web Worker architecture:

```
Main Thread                          Web Worker Thread
────────────────                     ────────────────────
DocumentRenderer                     WorkerDocumentHolder
  - render_sheet_from_bytes()          - load() [parse once]
  - render_page_from_bytes()           - parse_xlsx_sheet() [fast, cached]
  - caching layer                      - parse_docx_page()
       │                                    │
       └──── Transferable ◄────────────────┘
             ArrayBuffers (zero-copy)
```

This architecture:
- Parses documents in a background thread (non-blocking UI)
- Caches parsed document structure and styled sheet data
- Uses Transferable ArrayBuffers for zero-copy data transfer
- Supports request cancellation for rapid navigation

## Development

### Prerequisites

- Rust 1.75+
- LibreOffice (for core crate)
- pdfium library (for core crate)
- wasm-pack (for wasm crate)
- maturin (for python crate)

### Running Tests

```bash
# All Rust tests
cargo test

# Core crate only
cargo test --package office-to-png-core

# WASM crate only  
cargo test --package office-to-png-wasm

# With LibreOffice integration tests
LIBREOFFICE_PATH=/path/to/soffice cargo test
```

### Python Tests

```bash
cd crates/python

# Activate the virtual environment (if not already active)
source .venv/bin/activate  # On Windows: .venv\Scripts\activate

# Install test dependencies
pip install pytest pytest-asyncio

# Set library path for pdfium (run from repo root)
# macOS:
export DYLD_LIBRARY_PATH=$(pwd)/lib/lib:$DYLD_LIBRARY_PATH
# Linux:
export LD_LIBRARY_PATH=$(pwd)/lib/lib:$LD_LIBRARY_PATH

# Run all tests
pytest tests/ -v

# Run only tests that don't require LibreOffice
pytest tests/ -v -m "not requires_libreoffice"

# Run the quickstart example
python ../../examples/python_quickstart.py
```

### Building WASM

```bash
# Install wasm-pack
cargo install wasm-pack

# Build (standard Canvas 2D backend)
wasm-pack build crates/wasm --target web --release

# Build with WebGPU support
wasm-pack build crates/wasm --target web --release --features webgpu

# Output in crates/wasm/pkg/
```

### Running the Demo

```bash
# Build the WASM package first
wasm-pack build crates/wasm --target web

# Copy to demo folder
cp -r crates/wasm/pkg/* demo/pkg/

# Option 1: Local-only mode (WASM rendering in browser)
cd demo && python3 -m http.server 8000
# Open http://localhost:8000

# Option 2: With HD server mode (LibreOffice rendering)
# Terminal 1: Start the API server
PORT=8085 cargo run -p office-to-png-server

# Terminal 2: Serve the demo
cd demo && python3 -m http.server 8000
# Open http://localhost:8000, select "Server HD" mode
```

The demo includes:
- Drag-and-drop file upload (click or drop)
- Two rendering modes:
  - **Local (Fast)**: WASM-based rendering in browser
  - **Server HD**: LibreOffice-based rendering for 100% fidelity
- DOCX page navigation
- XLSX sheet tabs with smooth switching
- Zoom controls
- PNG export
- Web Worker integration for responsive UI

## Performance

Benchmarks on Apple M1 Pro (10 cores):

| Operation | Time |
|-----------|------|
| Single page DOCX | ~200ms |
| 10-page DOCX | ~800ms |
| Large XLSX (1000 rows) | ~500ms |
| Batch (100 files, pool=4) | ~15s |

WASM rendering (Chrome, Canvas 2D):

| Operation | Time |
|-----------|------|
| Parse DOCX | ~50ms |
| Render page | ~30ms |
| Export PNG | ~100ms |

WASM with Web Worker (recommended for interactive apps):

| Operation | Time | Notes |
|-----------|------|-------|
| Initial document parse | ~100-500ms | Once per document, in worker |
| Sheet switch (first time) | ~50-200ms | Extracts styling, cached |
| Sheet switch (cached) | ~5-20ms | Layout + draw only |
| Rapid sheet switching | Non-blocking | Debounced, stale results ignored |

## License

Licensed under either of:

- Apache License, Version 2.0 ([LICENSE-APACHE](LICENSE-APACHE) or http://www.apache.org/licenses/LICENSE-2.0)
- MIT license ([LICENSE-MIT](LICENSE-MIT) or http://opensource.org/licenses/MIT)

at your option.
