#!/usr/bin/env python3
"""
Single PDF to XLSX Converter Test Tool

A command-line utility to convert a single PDF file to XLSX format.
Useful for testing and converting individual files.

Usage:
    python test_single_xlsx.py <pdf_file>

Example:
    python test_single_xlsx.py mytest.pdf
    python test_single_xlsx.py pdf2excel.pdf
"""

import sys
import os
from pdf_to_xlsx_converter import (
    check_xlsx_dependencies,
    convert_pdf_to_xlsx,
    generate_xlsx_path,
    handle_xlsx_conversion_error,
    should_convert_file
)


def convert_single_pdf_to_xlsx(pdf_file: str) -> None:
    """
    Convert a single PDF file to XLSX format
    
    Args:
        pdf_file: Path to the PDF file to convert
    """
    try:
        print("Single PDF to XLSX Converter")
        print("=" * 50)
        print(f"Converting: {pdf_file}")
        print()
        
        # Check dependencies first
        if not check_xlsx_dependencies():
            print("\n❌ Cannot proceed due to missing dependencies")
            print("Please install the required dependencies and try again.")
            return
        
        print()
        
        # Check if the PDF file exists
        if not os.path.exists(pdf_file):
            print(f"❌ Error: PDF file '{pdf_file}' not found")
            print("Please check the file path and try again.")
            return
        
        if not os.path.isfile(pdf_file):
            print(f"❌ Error: '{pdf_file}' is not a file")
            return
        
        if not pdf_file.lower().endswith('.pdf'):
            print(f"❌ Error: '{pdf_file}' is not a PDF file")
            print("Please provide a file with .pdf extension.")
            return
        
        # Generate output path
        xlsx_path = generate_xlsx_path(pdf_file)
        
        print(f"📄 Input file: {pdf_file}")
        print(f"📊 Output file: {xlsx_path}")
        print()
        
        # Check if we should convert (handles overwrite protection)
        if not should_convert_file(pdf_file, xlsx_path):
            print("⏭️  Conversion cancelled by user")
            return
        
        # Perform the conversion
        print("🚀 Starting conversion...")
        success = convert_pdf_to_xlsx(pdf_file, xlsx_path)
        
        if success:
            print(f"\n🎉 Conversion completed successfully!")
            print(f"✅ Output saved to: {xlsx_path}")
            
            # Display file size information
            try:
                input_size = os.path.getsize(pdf_file)
                output_size = os.path.getsize(xlsx_path)
                print(f"📏 Input file size: {format_file_size(input_size)}")
                print(f"📏 Output file size: {format_file_size(output_size)}")
            except OSError:
                pass  # Ignore if we can't get file sizes
        else:
            print(f"\n❌ Conversion failed for unknown reason")
            
    except Exception as error:
        print(f"\n❌ Conversion failed:")
        handle_xlsx_conversion_error(error, os.path.basename(pdf_file))


def format_file_size(size_bytes: int) -> str:
    """
    Format file size in human-readable format
    
    Args:
        size_bytes: File size in bytes
        
    Returns:
        Formatted file size string
    """
    if size_bytes < 1024:
        return f"{size_bytes} bytes"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    elif size_bytes < 1024 * 1024 * 1024:
        return f"{size_bytes / (1024 * 1024):.1f} MB"
    else:
        return f"{size_bytes / (1024 * 1024 * 1024):.1f} GB"


def show_usage() -> None:
    """Display usage information"""
    print("Single PDF to XLSX Converter")
    print("=" * 50)
    print("Usage:")
    print("    python test_single_xlsx.py <pdf_file>")
    print()
    print("Examples:")
    print("    python test_single_xlsx.py mytest.pdf")
    print("    python test_single_xlsx.py pdf2excel.pdf")
    print("    python test_single_xlsx.py documents/report.pdf")
    print()
    print("The tool will:")
    print("• Extract tables from the PDF file")
    print("• Convert them to Excel format (.xlsx)")
    print("• Save the result in the same directory")
    print("• Ask for confirmation before overwriting existing files")


def main() -> None:
    """Main entry point for single file conversion"""
    try:
        # Check command line arguments
        if len(sys.argv) != 2:
            show_usage()
            return
        
        pdf_file = sys.argv[1]
        
        # Handle special help arguments
        if pdf_file.lower() in ['-h', '--help', 'help']:
            show_usage()
            return
        
        # Convert the file
        convert_single_pdf_to_xlsx(pdf_file)
        
    except KeyboardInterrupt:
        print("\n\n❌ Operation cancelled by user (Ctrl+C)")
        print("Conversion process terminated.")
        
    except Exception as error:
        print(f"\n❌ Unexpected error: {error}")
        print("Please check your input and try again.")
        
    finally:
        print("\nExiting Single PDF to XLSX Converter...")
        print("=" * 50)


if __name__ == "__main__":
    main()