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

- **Cross-platform**: Works on macOS, Linux, and Windows

## Crates

| Crate | Description |
|-------|-------------|
| `office-to-png-core` | Server-side conversion using LibreOffice + pdfium |
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
│   ├── wasm/           # Browser-side: Direct rendering
│   │   ├── docx_renderer.rs  # DOCX parsing + rendering
│   │   ├── xlsx_renderer.rs  # XLSX parsing + rendering
│   │   ├── renderer/
│   │   │   ├── traits.rs     # RenderBackend trait
│   │   │   ├── canvas2d.rs   # Canvas 2D implementation
│   │   │   └── webgpu.rs     # WebGPU implementation
│   │   └── text_shaper.rs    # cosmic-text integration
│   │
│   └── python/         # Python bindings via PyO3
│
├── tests/
│   └── fixtures/       # Test documents
│
└── examples/           # Usage examples
```

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
cd demo
python3 -m http.server 8080
# Open http://localhost:8080 in your browser
```

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

## License

Licensed under either of:

- Apache License, Version 2.0 ([LICENSE-APACHE](LICENSE-APACHE) or http://www.apache.org/licenses/LICENSE-2.0)
- MIT license ([LICENSE-MIT](LICENSE-MIT) or http://opensource.org/licenses/MIT)

at your option.
