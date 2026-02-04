"""
office-to-png: High-performance Office document to PNG conversion.

This library provides fast, parallelized conversion of Microsoft Office
documents (.docx, .xlsx) to PNG images using LibreOffice and pdfium.

Example:
    >>> import asyncio
    >>> from office_to_png import OfficeConverter
    >>>
    >>> async def main():
    ...     converter = OfficeConverter(pool_size=4, dpi=300)
    ...     result = await converter.convert("document.docx", "./output")
    ...     print(f"Rendered {result.page_count} pages")
    ...
    >>> asyncio.run(main())

For batch processing with progress:
    >>> async def batch_convert():
    ...     converter = OfficeConverter()
    ...     
    ...     def on_progress(p):
    ...         print(f"File {p.file_index + 1}/{p.total_files}: {p.current_file}")
    ...     
    ...     result = await converter.convert_batch(
    ...         ["doc1.docx", "doc2.xlsx"],
    ...         "./output",
    ...         progress_callback=on_progress
    ...     )
    ...     print(f"Converted {result.total_pages} pages total")
"""

from __future__ import annotations

# Import from the native Rust extension
from .office_to_png import (
    # Main converter class
    OfficeConverter,
    
    # Result types
    ConversionProgress,
    FileResult,
    BatchResult,
    PngPage,
    
    # Utility functions
    is_libreoffice_available,
    get_libreoffice_path,
    supported_extensions,
    is_supported_extension,
    init_logging,
)

__version__ = "0.1.0"
__all__ = [
    # Main class
    "OfficeConverter",
    
    # Result types
    "ConversionProgress",
    "FileResult", 
    "BatchResult",
    "PngPage",
    
    # Utilities
    "is_libreoffice_available",
    "get_libreoffice_path",
    "supported_extensions",
    "is_supported_extension",
    "init_logging",
]


def check_dependencies() -> dict[str, bool]:
    """Check if all dependencies are available.
    
    Returns:
        Dictionary with dependency names and their availability status.
    
    Example:
        >>> deps = check_dependencies()
        >>> if not deps['libreoffice']:
        ...     print("Please install LibreOffice")
    """
    return {
        "libreoffice": is_libreoffice_available(),
    }


def get_system_info() -> dict[str, str | None]:
    """Get system information relevant to office-to-png.
    
    Returns:
        Dictionary with system information.
    """
    import platform
    import os
    
    return {
        "platform": platform.system(),
        "platform_version": platform.version(),
        "python_version": platform.python_version(),
        "libreoffice_path": get_libreoffice_path(),
        "cpu_count": str(os.cpu_count()),
        "office_to_png_version": __version__,
    }
