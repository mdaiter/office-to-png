#!/usr/bin/env python3
"""
office-to-png Python Quickstart Example

Demonstrates the full workflow from installation to conversion.

Prerequisites:
    1. LibreOffice installed
    2. pdfium library available
    3. Python package built and installed

Installation:
    cd crates/python
    
    # Create and activate virtual environment
    python3 -m venv .venv
    source .venv/bin/activate  # On Windows: .venv\\Scripts\\activate
    
    # Install and build
    pip install maturin
    maturin develop
    
Usage:
    python examples/python_quickstart.py
"""

import asyncio
import sys
import time
from pathlib import Path


def main():
    """Run the quickstart example."""
    print("=" * 60)
    print("office-to-png Python Quickstart")
    print("=" * 60)
    print()
    
    # Step 1: Import and check dependencies
    print("Step 1: Checking dependencies...")
    try:
        from office_to_png import (
            OfficeConverter,
            is_libreoffice_available,
            get_libreoffice_path,
            supported_extensions,
            check_dependencies,
            get_system_info,
        )
        print("  [OK] office_to_png imported successfully")
    except ImportError as e:
        print(f"  [FAIL] Failed to import office_to_png: {e}")
        print()
        print("Please build and install the package first:")
        print("  cd crates/python")
        print("  pip install maturin")
        print("  maturin develop")
        sys.exit(1)
    
    # Step 2: Check system info
    print()
    print("Step 2: System Information")
    info = get_system_info()
    for key, value in info.items():
        print(f"  {key}: {value}")
    
    # Step 3: Check LibreOffice
    print()
    print("Step 3: LibreOffice Status")
    deps = check_dependencies()
    if deps["libreoffice"]:
        print(f"  [OK] LibreOffice found at: {get_libreoffice_path()}")
    else:
        print("  [FAIL] LibreOffice not found!")
        print("  Please install LibreOffice to continue.")
        sys.exit(1)
    
    print(f"  Supported formats: {', '.join(supported_extensions())}")
    
    # Step 4: Locate test fixtures
    print()
    print("Step 4: Locating test documents...")
    script_dir = Path(__file__).parent
    fixtures_dir = script_dir.parent / "tests" / "fixtures" / "output"
    
    if not fixtures_dir.exists():
        print(f"  [FAIL] Fixtures directory not found: {fixtures_dir}")
        print("  Please run from the repository root.")
        sys.exit(1)
    
    test_files = list(fixtures_dir.glob("*.docx")) + list(fixtures_dir.glob("*.xlsx"))
    print(f"  [OK] Found {len(test_files)} test documents")
    for f in test_files[:5]:  # Show first 5
        print(f"       - {f.name}")
    if len(test_files) > 5:
        print(f"       ... and {len(test_files) - 5} more")
    
    # Step 5: Run async conversions
    asyncio.run(run_conversions(fixtures_dir))
    
    print()
    print("=" * 60)
    print("Quickstart complete!")
    print("=" * 60)


async def run_conversions(fixtures_dir: Path):
    """Run the async conversion examples."""
    from office_to_png import OfficeConverter, ConversionProgress
    import tempfile
    
    # Create temporary output directory
    with tempfile.TemporaryDirectory() as output_dir:
        output_path = Path(output_dir)
        
        # Example 1: Single file conversion
        print()
        print("Step 5: Single File Conversion")
        print("-" * 40)
        
        converter = OfficeConverter(pool_size=2, dpi=150)
        print(f"  Created: {converter}")
        
        docx_file = fixtures_dir / "simple.docx"
        print(f"  Converting: {docx_file.name}")
        
        start = time.time()
        result = await converter.convert(str(docx_file), str(output_path))
        elapsed = time.time() - start
        
        print(f"  [OK] Converted {result.page_count} page(s) in {elapsed:.2f}s")
        print(f"  Output: {result.output_paths[0]}")
        print(f"  File size: {Path(result.output_paths[0]).stat().st_size:,} bytes")
        
        # Example 2: Batch conversion with progress
        print()
        print("Step 6: Batch Conversion with Progress")
        print("-" * 40)
        
        batch_files = [
            str(fixtures_dir / "simple.docx"),
            str(fixtures_dir / "formatted.docx"),
            str(fixtures_dir / "simple.xlsx"),
        ]
        print(f"  Converting {len(batch_files)} files...")
        
        progress_count = [0]
        
        def on_progress(p: ConversionProgress):
            progress_count[0] += 1
            print(f"    [{p.file_index + 1}/{p.total_files}] {p.stage}: {Path(p.current_file).name}")
        
        start = time.time()
        batch_result = await converter.convert_batch(
            batch_files,
            str(output_path),
            progress_callback=on_progress
        )
        elapsed = time.time() - start
        
        print()
        print(f"  [OK] Batch complete in {elapsed:.2f}s")
        print(f"  Successful: {batch_result.success_count}")
        print(f"  Failed: {batch_result.failure_count}")
        print(f"  Total pages: {batch_result.total_pages}")
        
        # Example 3: Custom DPI (using per-request DPI override)
        print()
        print("Step 7: Custom DPI Example")
        print("-" * 40)
        
        print(f"  Using existing converter with DPI override (300)")
        
        start = time.time()
        hd_result = await converter.convert(
            str(fixtures_dir / "simple.docx"),
            str(output_path),
            dpi=300,  # Override DPI for this request
            output_prefix="high_dpi"
        )
        elapsed = time.time() - start
        
        hd_size = Path(hd_result.output_paths[0]).stat().st_size
        print(f"  [OK] 300 DPI conversion in {elapsed:.2f}s")
        print(f"  Output size: {hd_size:,} bytes (larger due to higher DPI)")
        
        # Example 4: Check converter health
        print()
        print("Step 8: Converter Health")
        print("-" * 40)
        
        health = await converter.health()
        print(f"  Pool status: {health}")
        
        # Cleanup
        print()
        print("Step 9: Cleanup")
        print("-" * 40)
        await converter.shutdown()
        print("  [OK] Converter shut down")
        
        # Summary
        print()
        print("Summary")
        print("-" * 40)
        all_pngs = list(output_path.glob("*.png"))
        total_size = sum(p.stat().st_size for p in all_pngs)
        print(f"  Total PNG files generated: {len(all_pngs)}")
        print(f"  Total size: {total_size:,} bytes ({total_size / 1024 / 1024:.2f} MB)")


if __name__ == "__main__":
    main()
