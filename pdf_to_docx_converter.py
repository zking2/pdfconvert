#!/usr/bin/env python3
"""
PDF to DOCX Converter

A command-line utility to convert PDF files to DOCX format using the pdf2docx library.
Scans the current working directory for PDF files and converts them to DOCX format.
"""

import os
from dataclasses import dataclass
from typing import List, Optional
from pdf2docx import Converter


@dataclass
class ConversionResult:
    """Data class to track individual file conversion status"""
    file_name: str
    status: str  # 'success', 'failed', 'skipped'
    error_message: Optional[str] = None


@dataclass
class ConversionSummary:
    """Data class to track overall batch conversion results"""
    total_files: int
    successful: int
    failed: int
    skipped: int
    results: List[ConversionResult]


# Core function interfaces with type hints
def get_pdf_files(directory_path: str) -> List[str]:
    """
    Scan directory and identify PDF files
    
    Args:
        directory_path: Path to the directory to scan for PDF files
        
    Returns:
        List of PDF file paths found in the directory
        
    Raises:
        FileNotFoundError: If the directory doesn't exist
        PermissionError: If the directory is not accessible
    """
    if not os.path.exists(directory_path):
        raise FileNotFoundError(f"Directory not found: {directory_path}")
    
    if not os.path.isdir(directory_path):
        raise NotADirectoryError(f"Path is not a directory: {directory_path}")
    
    try:
        pdf_files = []
        for file_name in os.listdir(directory_path):
            file_path = os.path.join(directory_path, file_name)
            
            # Check if it's a file (not a directory) and has .pdf extension
            if os.path.isfile(file_path) and file_name.lower().endswith('.pdf'):
                pdf_files.append(file_path)
        
        return sorted(pdf_files)  # Return sorted list for consistent ordering
        
    except PermissionError as e:
        raise PermissionError(f"Permission denied accessing directory: {directory_path}") from e


def is_valid_pdf_file(pdf_path: str) -> tuple[bool, str]:
    """
    Check if a file is a valid PDF by examining its structure
    
    Args:
        pdf_path: Path to the PDF file to validate
        
    Returns:
        Tuple of (is_valid, error_message)
        - is_valid: True if the file appears to be a valid PDF
        - error_message: Description of the issue if not valid, empty string if valid
    """
    try:
        with open(pdf_path, 'rb') as f:
            # Read first 1024 bytes to check PDF structure
            header = f.read(1024)
            
            # Check for PDF header
            if not header.startswith(b'%PDF'):
                return False, "File does not have a valid PDF header"
            
            # Check file size first - empty or very small files are likely invalid
            f.seek(0, 2)  # Seek to end
            file_size = f.tell()
            if file_size < 100:  # PDF files should be at least 100 bytes
                return False, "File is too small to be a valid PDF"
            
            # Check for basic PDF structure markers
            if b'obj' not in header and b'endobj' not in header:
                # Read more of the file to look for object markers
                f.seek(0)
                content = f.read(4096)
                if b'obj' not in content:
                    return False, "File does not contain valid PDF object structure"
            
            # Look for EOF marker near the end of file
            f.seek(max(0, file_size - 1024))
            tail = f.read()
            if b'%%EOF' not in tail:
                return False, "File does not have a valid PDF end marker"
            
        return True, ""
        
    except Exception as e:
        return False, f"Error reading file: {e}"


def validate_file_accessibility(pdf_path: str, docx_path: str) -> None:
    """
    Validate that the PDF file is accessible and the output location is writable
    
    Args:
        pdf_path: Path to the input PDF file
        docx_path: Path where the output DOCX file should be saved
        
    Raises:
        FileNotFoundError: If the PDF file doesn't exist or is not accessible
        PermissionError: If there are permission issues with file access
        ValueError: If the PDF file is not valid
        OSError: If there are other file system issues
    """
    # Check if PDF file exists and is accessible
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"PDF file not found: {pdf_path}")
    
    if not os.path.isfile(pdf_path):
        raise FileNotFoundError(f"Path is not a file: {pdf_path}")
    
    # Check if PDF file is readable and valid
    try:
        # Validate PDF file structure
        is_valid, error_msg = is_valid_pdf_file(pdf_path)
        if not is_valid:
            raise ValueError(f"Invalid PDF file: {error_msg}")
            
    except PermissionError:
        raise PermissionError(f"Permission denied reading PDF file: {pdf_path}")
    except OSError as e:
        raise OSError(f"Cannot access PDF file: {pdf_path} - {e}")
    except ValueError:
        # Re-raise ValueError as-is
        raise
    except Exception as e:
        raise OSError(f"Error validating PDF file: {pdf_path} - {e}")
    
    # Check if output directory exists and is writable
    output_dir = os.path.dirname(docx_path)
    if output_dir and not os.path.exists(output_dir):
        raise FileNotFoundError(f"Output directory does not exist: {output_dir}")
    
    # Test write permissions by creating a temporary file
    try:
        temp_file = docx_path + ".tmp"
        with open(temp_file, 'w') as f:
            f.write("test")
        os.remove(temp_file)
    except PermissionError:
        raise PermissionError(f"Permission denied writing to output location: {docx_path}")
    except OSError as e:
        raise OSError(f"Cannot write to output location: {docx_path} - {e}")


def convert_pdf_to_docx(pdf_path: str, docx_path: str) -> bool:
    """
    Handle individual PDF to DOCX conversion using pdf2docx Converter class
    
    Args:
        pdf_path: Path to the input PDF file
        docx_path: Path where the output DOCX file should be saved
        
    Returns:
        True if conversion was successful, False otherwise
        
    Raises:
        FileNotFoundError: If the PDF file doesn't exist
        PermissionError: If there are permission issues with file access
        ValueError: If the file is not a valid PDF
        Exception: For pdf2docx conversion errors
    """
    # Validate file accessibility before attempting conversion
    validate_file_accessibility(pdf_path, docx_path)
    
    # Initialize converter with None to ensure proper cleanup
    cv = None
    try:
        # Create converter instance with enhanced error handling
        try:
            cv = Converter(pdf_path)
        except Exception as e:
            # Enhance error message for converter creation failures
            if "password" in str(e).lower() or "encrypted" in str(e).lower():
                raise Exception(f"PDF is password-protected: {e}")
            elif "corrupt" in str(e).lower() or "invalid" in str(e).lower():
                raise Exception(f"PDF file appears to be corrupted: {e}")
            else:
                raise Exception(f"Failed to initialize PDF converter: {e}")
        
        # Perform the conversion with timeout protection
        try:
            cv.convert(docx_path)
        except Exception as e:
            # Enhance error messages for conversion failures
            error_msg = str(e).lower()
            if "memory" in error_msg:
                raise MemoryError(f"Insufficient memory to convert PDF: {e}")
            elif "timeout" in error_msg:
                raise Exception(f"Conversion timed out - PDF may be too complex: {e}")
            elif "unsupported" in error_msg:
                raise Exception(f"PDF contains unsupported features: {e}")
            else:
                raise e
        
        # Verify the output file was created and is valid
        if not os.path.exists(docx_path):
            raise Exception(f"Conversion completed but output file not found: {docx_path}")
        
        # Check if output file has reasonable size (not empty)
        if os.path.getsize(docx_path) == 0:
            raise Exception(f"Conversion produced empty output file: {docx_path}")
        
        return True
        
    except Exception as e:
        # Clean up partial output file if it exists
        if os.path.exists(docx_path):
            try:
                os.remove(docx_path)
            except OSError:
                pass  # Ignore cleanup errors
        
        # Re-raise the original exception
        raise e
        
    finally:
        # Ensure proper resource cleanup
        if cv is not None:
            try:
                cv.close()
            except Exception:
                pass  # Ignore cleanup errors


def check_file_exists(file_path: str) -> bool:
    """
    Check if a file exists at the given path
    
    Args:
        file_path: Path to check for file existence
        
    Returns:
        True if file exists and is a regular file, False otherwise
    """
    return os.path.exists(file_path) and os.path.isfile(file_path)


def generate_docx_path(pdf_path: str) -> str:
    """
    Generate output DOCX file path from input PDF file path
    
    Args:
        pdf_path: Path to the input PDF file
        
    Returns:
        Path for the output DOCX file (same directory, same name, .docx extension)
    """
    # Get the directory and filename without extension
    directory = os.path.dirname(pdf_path)
    filename_without_ext = os.path.splitext(os.path.basename(pdf_path))[0]
    
    # Create the DOCX path
    docx_filename = f"{filename_without_ext}.docx"
    docx_path = os.path.join(directory, docx_filename)
    
    return docx_path


def handle_conversion_error(error: Exception, file_name: str) -> None:
    """
    Manage exceptions and provide user-friendly error messages with specific guidance
    
    Args:
        error: The exception that occurred during conversion
        file_name: Name of the file that caused the error
    """
    error_type = type(error).__name__
    error_message = str(error).lower()
    
    if isinstance(error, ImportError):
        print(f"❌ Import Error for {file_name}: {error}")
        print("   Please install the required pdf2docx library: pip install pdf2docx")
    elif isinstance(error, FileNotFoundError):
        print(f"❌ File Not Found for {file_name}: {error}")
        print("   The PDF file may have been moved or deleted during processing")
    elif isinstance(error, PermissionError):
        print(f"❌ Permission Error for {file_name}: {error}")
        print("   Check file permissions and ensure you have read/write access")
        print("   Try running with elevated permissions or check file ownership")
    elif isinstance(error, OSError):
        if "disk" in error_message or "space" in error_message:
            print(f"❌ Disk Space Error for {file_name}: {error}")
            print("   Insufficient disk space to complete the conversion")
            print("   Free up disk space and try again")
        elif "memory" in error_message:
            print(f"❌ Memory Error for {file_name}: {error}")
            print("   Insufficient memory to process this PDF file")
            print("   Try closing other applications or processing smaller files")
        else:
            print(f"❌ System Error for {file_name}: {error}")
            print("   A system-level error occurred during file processing")
    elif hasattr(error, '__module__') and 'pdf2docx' in str(error.__module__):
        if "password" in error_message or "encrypted" in error_message:
            print(f"❌ Password Protected PDF for {file_name}: {error}")
            print("   This PDF is password-protected and cannot be converted")
            print("   Remove password protection or use a different file")
        elif "corrupt" in error_message or "invalid" in error_message or "damaged" in error_message:
            print(f"❌ Corrupted PDF for {file_name}: {error}")
            print("   The PDF file appears to be corrupted or damaged")
            print("   Try opening the file in a PDF viewer to verify its integrity")
        elif "unsupported" in error_message or "format" in error_message:
            print(f"❌ Unsupported PDF Format for {file_name}: {error}")
            print("   This PDF format is not supported by the converter")
            print("   Try converting the PDF to a standard format first")
        else:
            print(f"❌ PDF Conversion Error for {file_name}: {error}")
            print("   The PDF file may be corrupted, password-protected, or in an unsupported format")
    elif isinstance(error, MemoryError):
        print(f"❌ Memory Error for {file_name}: {error}")
        print("   The PDF file is too large to process with available memory")
        print("   Try processing smaller files or increase available memory")
    elif isinstance(error, UnicodeError):
        print(f"❌ Encoding Error for {file_name}: {error}")
        print("   The PDF contains characters that cannot be properly encoded")
        print("   The file may contain special fonts or non-standard text encoding")
    else:
        print(f"❌ Unexpected Error for {file_name} ({error_type}): {error}")
        print("   An unexpected error occurred during conversion")
        print("   Please check the PDF file integrity and try again")


def check_dependencies() -> bool:
    """
    Verify pdf2docx installation and other dependencies
    
    Returns:
        True if all dependencies are available, False otherwise
    """
    return _check_dependencies_impl()


def _check_dependencies_impl() -> bool:
    """
    Internal implementation of dependency checking for easier testing
    
    Returns:
        True if all dependencies are available, False otherwise
    """
    try:
        # Try to import pdf2docx to verify it's installed
        import pdf2docx
        from pdf2docx import Converter
        
        # Check if we can access the main Converter class
        if not hasattr(pdf2docx, 'Converter'):
            print("❌ Error: pdf2docx library is installed but Converter class is not available")
            return False
            
        print("✅ Dependencies check passed: pdf2docx is available")
        return True
        
    except ImportError as e:
        print("❌ Dependency Error: pdf2docx library is not installed")
        print("   Please install it using: pip install pdf2docx")
        print(f"   Error details: {e}")
        return False
    except Exception as e:
        print(f"❌ Unexpected error during dependency check: {e}")
        return False


def display_progress(current: int, total: int, file_name: str) -> None:
    """
    Display current file being processed with progress information
    
    Args:
        current: Current file number being processed (1-based)
        total: Total number of files to process
        file_name: Name of the current file being processed
    """
    if total <= 0:
        print(f"🔄 Processing: {file_name}")
    else:
        percentage = (current / total) * 100
        print(f"🔄 Processing ({current}/{total} - {percentage:.0f}%): {file_name}")


def display_file_count(count: int) -> None:
    """
    Display the total number of PDF files found
    
    Args:
        count: Number of PDF files found in the directory
    """
    if count == 0:
        print("📁 No PDF files found in the current directory")
    elif count == 1:
        print("📁 Found 1 PDF file to convert")
    else:
        print(f"📁 Found {count} PDF files to convert")


def prompt_overwrite(file_name: str) -> bool:
    """
    Prompt user for confirmation on file overwrite
    
    Args:
        file_name: Name of the DOCX file that already exists
        
    Returns:
        True if user confirms overwrite, False if user declines
    """
    while True:
        try:
            response = input(f"⚠️  File '{file_name}' already exists. Overwrite? (y/n): ").strip().lower()
            
            if response in ['y', 'yes']:
                return True
            elif response in ['n', 'no']:
                return False
            else:
                print("Please enter 'y' for yes or 'n' for no.")
                
        except (EOFError, KeyboardInterrupt):
            # Handle Ctrl+C or EOF gracefully
            print("\n❌ Operation cancelled by user")
            return False


def should_convert_file(pdf_path: str, docx_path: str) -> bool:
    """
    Check if a file should be converted, handling overwrite protection
    
    Args:
        pdf_path: Path to the input PDF file
        docx_path: Path where the output DOCX file would be saved
        
    Returns:
        True if the file should be converted, False if it should be skipped
    """
    # If the DOCX file doesn't exist, proceed with conversion
    if not check_file_exists(docx_path):
        return True
    
    # If the DOCX file exists, prompt user for overwrite confirmation
    docx_filename = os.path.basename(docx_path)
    return prompt_overwrite(docx_filename)


def display_summary(successful: int, failed: int, skipped: int) -> None:
    """
    Display final conversion results summary
    
    Args:
        successful: Number of files successfully converted
        failed: Number of files that failed to convert
        skipped: Number of files that were skipped (due to existing DOCX files)
    """
    total = successful + failed + skipped
    
    print("\n" + "=" * 50)
    print("📊 CONVERSION SUMMARY")
    print("=" * 50)
    
    if total == 0:
        print("No files were processed")
        return
    
    print(f"Total files processed: {total}")
    
    if successful > 0:
        print(f"✅ Successfully converted: {successful}")
    
    if failed > 0:
        print(f"❌ Failed conversions: {failed}")
    
    if skipped > 0:
        print(f"⏭️  Skipped files: {skipped}")
    
    # Calculate and display success rate
    if total > 0:
        success_rate = (successful / total) * 100
        print(f"📈 Success rate: {success_rate:.1f}%")
    
    print("=" * 50)


def display_conversion_success(file_name: str) -> None:
    """
    Display success message for individual file conversion
    
    Args:
        file_name: Name of the file that was successfully converted
    """
    print(f"✅ Successfully converted: {file_name}")


def display_conversion_skipped(file_name: str) -> None:
    """
    Display message when a file is skipped due to existing DOCX
    
    Args:
        file_name: Name of the file that was skipped
    """
    print(f"⏭️  Skipped (file exists): {file_name}")


def pdf_to_docx_batch_convert() -> None:
    """
    Orchestrate the entire batch conversion process
    
    This function integrates file discovery, conversion, error handling, and progress feedback.
    It continues processing after individual file failures and tracks conversion statistics.
    
    Requirements addressed:
    - 1.1: Scan current working directory for PDF files and convert to DOCX
    - 2.1: Handle errors gracefully and continue with next file
    - 3.4: Display summary of successful and failed conversions
    """
    # Initialize conversion statistics
    successful_count = 0
    failed_count = 0
    skipped_count = 0
    conversion_results = []
    
    try:
        # Get current working directory
        current_directory = os.getcwd()
        
        # Discover PDF files in the current directory
        pdf_files = get_pdf_files(current_directory)
        
        # Display total number of files found
        display_file_count(len(pdf_files))
        
        # If no PDF files found, exit early
        if not pdf_files:
            display_summary(successful_count, failed_count, skipped_count)
            return
        
        print()  # Add blank line for better readability
        
        # Process each PDF file
        for index, pdf_path in enumerate(pdf_files, 1):
            pdf_filename = os.path.basename(pdf_path)
            docx_path = generate_docx_path(pdf_path)
            docx_filename = os.path.basename(docx_path)
            
            # Display progress
            display_progress(index, len(pdf_files), pdf_filename)
            
            try:
                # Check if we should convert this file (handles overwrite protection)
                if not should_convert_file(pdf_path, docx_path):
                    # User declined to overwrite existing file
                    display_conversion_skipped(pdf_filename)
                    skipped_count += 1
                    conversion_results.append(ConversionResult(
                        file_name=pdf_filename,
                        status='skipped',
                        error_message='User declined to overwrite existing DOCX file'
                    ))
                    continue
                
                # Attempt the conversion
                success = convert_pdf_to_docx(pdf_path, docx_path)
                
                if success:
                    display_conversion_success(pdf_filename)
                    successful_count += 1
                    conversion_results.append(ConversionResult(
                        file_name=pdf_filename,
                        status='success'
                    ))
                else:
                    # This shouldn't happen with current implementation, but handle it
                    print(f"❌ Conversion failed for {pdf_filename}: Unknown error")
                    failed_count += 1
                    conversion_results.append(ConversionResult(
                        file_name=pdf_filename,
                        status='failed',
                        error_message='Unknown conversion error'
                    ))
                
            except Exception as error:
                # Handle conversion error gracefully and continue with next file
                handle_conversion_error(error, pdf_filename)
                failed_count += 1
                conversion_results.append(ConversionResult(
                    file_name=pdf_filename,
                    status='failed',
                    error_message=str(error)
                ))
                # Continue processing the next file
                continue
        
        # Display final summary
        display_summary(successful_count, failed_count, skipped_count)
        
    except Exception as error:
        # Handle errors in file discovery or other critical failures
        print(f"\n❌ Critical error during batch conversion: {error}")
        print("Batch conversion process terminated.")
        
        # Still display summary of any files that were processed
        if successful_count > 0 or failed_count > 0 or skipped_count > 0:
            display_summary(successful_count, failed_count, skipped_count)


def main() -> None:
    """
    Main entry point for the PDF to DOCX converter application
    
    This function serves as the application entry point with proper error handling,
    dependency checking at startup, and final summary display.
    
    Requirements addressed:
    - 1.4: Provide feedback on number of files processed
    - 2.3: Handle missing pdf2docx library with installation instructions
    - 3.4: Display summary of successful and failed conversions
    """
    try:
        print("PDF to DOCX Converter")
        print("=" * 50)
        print("Converting PDF files in current directory to DOCX format...")
        print()
        
        # Check dependencies before proceeding
        if not check_dependencies():
            print("\n❌ Cannot proceed due to missing dependencies")
            print("Please install the required dependencies and try again.")
            return
        
        print()  # Add blank line after dependency check
        
        # Execute the main batch conversion process
        pdf_to_docx_batch_convert()
        
        print("\n🎉 Conversion process completed!")
        
    except KeyboardInterrupt:
        # Handle Ctrl+C gracefully
        print("\n\n❌ Operation cancelled by user (Ctrl+C)")
        print("Conversion process terminated.")
        
    except Exception as error:
        # Handle any unexpected errors at the application level
        print(f"\n❌ Unexpected application error: {error}")
        print("Please check your environment and try again.")
        print("If the problem persists, please report this issue.")
        
    finally:
        # Ensure clean exit
        print("\nExiting PDF to DOCX Converter...")
        print("=" * 50)


if __name__ == "__main__":
    main()