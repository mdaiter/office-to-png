#!/usr/bin/env python3
"""
Example: Simple single-file conversion.

This is the minimal example for converting a single document.
"""

import asyncio
import sys


async def main():
    if len(sys.argv) < 3:
        print("Usage: python simple_convert.py <input_file> <output_dir>")
        print("Example: python simple_convert.py document.docx ./output")
        return 1

    input_file = sys.argv[1]
    output_dir = sys.argv[2]

    try:
        from office_to_png import OfficeConverter
    except ImportError:
        print("Error: office_to_png not installed")
        return 1

    # Create converter with defaults
    converter = OfficeConverter()

    # Convert the file
    print(f"Converting {input_file}...")
    result = await converter.convert(input_file, output_dir)

    # Print results
    print(f"Success! Generated {result.page_count} pages in {result.duration_secs:.2f}s")
    for path in result.output_paths:
        print(f"  - {path}")

    # Cleanup
    await converter.shutdown()
    return 0


if __name__ == "__main__":
    sys.exit(asyncio.run(main()))
