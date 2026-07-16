import os
import sys
import argparse
from pathlib import Path
try:
    from pdf2image import convert_from_path
except ImportError:
    print("Error: pdf2image library is not installed. Please install it using 'pip install pdf2image'.")
    sys.exit(1)

def batch_convert(directory, poppler_path=None):
    directory = Path(directory)
    if not directory.exists() or not directory.is_dir():
        print(f"Error: Invalid directory '{directory}'")
        return

    pdf_files = list(directory.glob("*.pdf"))
    if not pdf_files:
        print(f"No PDF files found in {directory}")
        return

    print(f"Found {len(pdf_files)} PDF files in {directory}...")

    # Resolve poppler_path: explicit arg/env wins, then on-demand cache
    # (auto-downloads on Windows), finally let pdf2image try PATH.
    poppler_context = ""  # reason for the install hint if conversion fails
    if not poppler_path:
        try:
            from backend.resource_manager import get_poppler_bin_path
            cached_bin = get_poppler_bin_path()
            if cached_bin:
                poppler_path = str(cached_bin)
                print(f"Using cached Poppler from: {poppler_path}")
            else:
                from backend.dependency_hints import poppler_failure_context
                poppler_context = poppler_failure_context()
        except Exception as e:
            print(f"Poppler cache lookup failed: {e}")

    success_count = 0
    fail_count = 0
    poppler_unavailable = False

    for pdf_file in pdf_files:
        try:
            print(f"Converting {pdf_file.name}...")
            # Convert first page only (assuming single page figures)
            images = convert_from_path(str(pdf_file), first_page=1, last_page=1, poppler_path=poppler_path)

            if images:
                output_file = pdf_file.with_suffix('.png')
                images[0].save(output_file, 'PNG')
                print(f"  -> Saved to {output_file.name}")
                success_count += 1

                # Delete original PDF if successful
                try:
                    pdf_file.unlink()
                    print(f"  -> Deleted original PDF: {pdf_file.name}")
                except Exception as del_e:
                    print(f"  -> Warning: Failed to delete PDF {pdf_file.name}: {del_e}")

            else:
                print(f"  -> Warning: No images extracted from {pdf_file.name}")
                fail_count += 1

        except Exception as e:
            # If Poppler is missing, every remaining PDF will fail the same way —
            # surface one actionable platform-specific hint and abort instead of
            # spamming the same error per file.
            try:
                from backend.dependency_hints import is_poppler_missing_error, poppler_install_hint
                if is_poppler_missing_error(e):
                    print(f"  -> Failed: {pdf_file.name}")
                    print(poppler_install_hint(context=poppler_context))
                    remaining = len(pdf_files) - success_count - fail_count - 1
                    if remaining > 0:
                        print(f"\nAborting batch: Poppler unavailable; {remaining} more PDF(s) skipped.")
                        fail_count += remaining
                    fail_count += 1
                    poppler_unavailable = True
                    break
            except Exception:
                pass
            print(f"  -> Failed: {e}")
            fail_count += 1

    print(f"\nConversion Complete.")
    print(f"Success: {success_count}")
    print(f"Failed: {fail_count}")

    if fail_count > 0 and not poppler_unavailable:
        print("\nTroubleshooting:")
        try:
            from backend.dependency_hints import poppler_install_hint
            print(poppler_install_hint())
        except Exception:
            print("Ensure 'poppler' is installed and available on PATH.")
        
    # Exit with error if no files were successfully converted
    if success_count == 0 and len(pdf_files) > 0:
        sys.exit(1)
    # Exit with error if partial failure (optional, maybe just warning?)
    # Let's stick to error if ALL failed, otherwise 0.
    
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Batch convert single-page PDFs to PNG images.")
    parser.add_argument("directory", help="Directory containing PDF files")
    parser.add_argument("--poppler-path", help="Path to poppler bin folder (optional)", default=None)
    
    args = parser.parse_args()
    
    # Try to detect poppler path from env var if not provided
    poppler_path = args.poppler_path or os.environ.get('POPPLER_PATH')
    
    batch_convert(args.directory, poppler_path)
