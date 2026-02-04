#!/usr/bin/env python3
"""
Example: Batch convert Office documents to PNG.

This example demonstrates:
- Creating a converter with custom settings
- Batch converting multiple files
- Progress tracking with callbacks
- Error handling

Usage:
    python batch_convert.py input_dir output_dir [--dpi 300] [--workers 4]
"""

import asyncio
import argparse
import sys
from pathlib import Path
from typing import Optional


def create_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Convert Office documents to PNG images"
    )
    parser.add_argument(
        "input_dir",
        type=Path,
        help="Directory containing input documents",
    )
    parser.add_argument(
        "output_dir",
        type=Path,
        help="Directory for output PNG files",
    )
    parser.add_argument(
        "--dpi",
        type=int,
        default=300,
        help="Output DPI (default: 300)",
    )
    parser.add_argument(
        "--workers",
        type=int,
        default=None,
        help="Number of LibreOffice workers (default: CPU count)",
    )
    parser.add_argument(
        "--extensions",
        type=str,
        default="docx,xlsx,doc,xls",
        help="File extensions to process (comma-separated)",
    )
    parser.add_argument(
        "--verbose",
        "-v",
        action="store_true",
        help="Verbose output",
    )
    return parser


async def main(args: argparse.Namespace) -> int:
    try:
        from office_to_png import (
            OfficeConverter,
            is_libreoffice_available,
            supported_extensions,
        )
    except ImportError:
        print("Error: office_to_png not installed")
        print("Install with: pip install office-to-png")
        return 1

    # Check dependencies
    if not is_libreoffice_available():
        print("Error: LibreOffice not found")
        print("Install LibreOffice and ensure 'soffice' is in PATH")
        return 1

    # Validate input directory
    if not args.input_dir.is_dir():
        print(f"Error: Input directory not found: {args.input_dir}")
        return 1

    # Create output directory
    args.output_dir.mkdir(parents=True, exist_ok=True)

    # Find files to convert
    extensions = [ext.strip().lower() for ext in args.extensions.split(",")]
    valid_extensions = set(supported_extensions())
    
    input_files = []
    for ext in extensions:
        if ext not in valid_extensions:
            print(f"Warning: Unsupported extension '{ext}', skipping")
            continue
        input_files.extend(args.input_dir.glob(f"*.{ext}"))
        input_files.extend(args.input_dir.glob(f"*.{ext.upper()}"))

    if not input_files:
        print(f"No files found in {args.input_dir}")
        return 0

    print(f"Found {len(input_files)} files to convert")
    
    # Create converter
    print(f"Initializing converter (workers={args.workers or 'auto'}, dpi={args.dpi})")
    converter = OfficeConverter(
        pool_size=args.workers,
        dpi=args.dpi,
    )

    # Progress callback
    def on_progress(progress):
        pct = (progress.file_index + 1) / progress.total_files * 100
        pages = progress.pages_completed
        total = progress.total_pages if progress.total_pages else "?"
        stage = progress.stage
        
        if args.verbose:
            print(
                f"[{pct:5.1f}%] {progress.current_file}: "
                f"{pages}/{total} pages ({stage})"
            )
        else:
            # Simple progress bar
            bar_len = 40
            filled = int(bar_len * (progress.file_index + 1) / progress.total_files)
            bar = "█" * filled + "░" * (bar_len - filled)
            print(f"\r[{bar}] {pct:.1f}% - {progress.current_file}", end="", flush=True)

    # Convert files
    print("Starting conversion...")
    
    result = await converter.convert_batch(
        input_paths=[str(f) for f in input_files],
        output_dir=str(args.output_dir),
        progress_callback=on_progress,
    )

    # Print newline after progress bar
    if not args.verbose:
        print()

    # Print results
    print("\n" + "=" * 60)
    print("Conversion Complete!")
    print("=" * 60)
    print(f"  Successful: {result.success_count}")
    print(f"  Failed:     {result.failure_count}")
    print(f"  Total pages: {result.total_pages}")
    print(f"  Duration:    {result.total_duration_secs:.2f}s")
    
    if result.total_pages > 0:
        pages_per_sec = result.total_pages / result.total_duration_secs
        print(f"  Throughput:  {pages_per_sec:.1f} pages/sec")

    # Print failures
    if result.failed:
        print("\nFailed files:")
        for path, error in result.failed:
            print(f"  - {path}: {error}")

    # Print output locations
    if result.successful and args.verbose:
        print("\nOutput files:")
        for file_result in result.successful[:5]:
            for output_path in file_result.output_paths[:2]:
                print(f"  - {output_path}")
            if len(file_result.output_paths) > 2:
                print(f"    ... and {len(file_result.output_paths) - 2} more pages")
        if len(result.successful) > 5:
            print(f"  ... and {len(result.successful) - 5} more files")

    # Shutdown
    await converter.shutdown()

    return 0 if result.all_succeeded else 1


if __name__ == "__main__":
    parser = create_parser()
    args = parser.parse_args()
    
    try:
        exit_code = asyncio.run(main(args))
        sys.exit(exit_code)
    except KeyboardInterrupt:
        print("\nCancelled by user")
        sys.exit(130)
