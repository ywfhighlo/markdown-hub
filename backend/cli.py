#!/usr/bin/env python3
"""
Office & Docs Converter - CLI entry point.
Unified entry for the VS Code extension frontend; uses a factory pattern
to select and instantiate the correct converter.
"""

import sys
import os

# Inject the vendor path first so bundled libraries take precedence.
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'vendor'))

import argparse
import json
import logging
import base64
from typing import Dict, Type

# Add the current directory to the path so the converters package is importable.
sys.path.insert(0, os.path.dirname(__file__))

from converters.base_converter import BaseConverter
# Import the converter registry from __init__.py
from converters import CONVERTER_REGISTRY

def setup_logging():
    """Configure logging."""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[logging.StreamHandler(sys.stderr)]
    )

def get_converter(conversion_type: str, output_dir: str, **kwargs) -> BaseConverter:
    """Converter factory"""
    if conversion_type not in CONVERTER_REGISTRY:
        raise ValueError(f"Unsupported conversion type: {conversion_type}")

    converter_class = CONVERTER_REGISTRY[conversion_type]

    # Pass the output format to MdToOfficeConverter.
    if conversion_type.startswith('md-to-'):
        output_format = conversion_type.split('-')[-1]  # extract docx/pdf/html
        kwargs['output_format'] = output_format

    return converter_class(output_dir, **kwargs)

def main():
    parser = argparse.ArgumentParser(
        description="Markdown Hub - document conversion tool",
        formatter_class=argparse.RawTextHelpFormatter
    )
    parser.add_argument('--conversion-type', required=True, 
                       choices=list(CONVERTER_REGISTRY.keys()),
                       help='Conversion type')
    parser.add_argument('--input-path', required=True,
                       help='Input file or directory path')
    parser.add_argument('--output-dir', required=True,
                       help='Output directory')
    parser.add_argument('--docx-template-path',
                       help='Optional DOCX template file path')
    parser.add_argument('--pptx-template-path',
                       help='Optional PPTX template file path')
    parser.add_argument('--project-name', help='Project name (optional)')
    parser.add_argument('--author', help='Author name (optional)')
    parser.add_argument('--mobilephone', help='Contact phone (optional)')
    parser.add_argument('--email', help='Email address (optional)')
    parser.add_argument('--promote-headings', action='store_true',
                       help='Promote Markdown heading levels by one (e.g. ## -> Heading 1)')
    parser.add_argument('--code-highlight-theme', default='pygments',
                       help='pandoc code block highlight theme (pygments/tango/espresso/zenburn/kate/monochrome/breezedark/haddock/off, default: pygments)')
    parser.add_argument('--verbose', '-v', action='store_true',
                       help='Enable verbose log output')
    parser.add_argument('--poppler-path',
                       help='Path to Poppler tools (for PDF OCR)')
    parser.add_argument('--tesseract-cmd',
                       help='Tesseract-OCR command or path (for PDF OCR)')
    # SVG conversion parameters
    parser.add_argument('--svg-dpi', type=int, default=300,
                       help='DPI for SVG to PNG (default: 300)')
    parser.add_argument('--svg-output-width', type=int, default=800,
                       help='Output width for SVG to PNG (default: 800px)')
    
    args = parser.parse_args()
    
    # Set log level
    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)
    else:
        setup_logging()

    # Build the progress reporter
    def report_progress(stage: str, percentage: int = None):
        progress = {
            "type": "progress",
            "stage": stage
        }
        if percentage is not None:
            progress["percentage"] = percentage

        # Base64-encode so UTF-8 content passes through stdout safely.
        json_str = json.dumps(progress, ensure_ascii=False)
        encoded_str = base64.b64encode(json_str.encode('utf-8')).decode('ascii')
        print(encoded_str, flush=True)

    try:
        report_progress("preparing")

        # Resolve the converter class
        converter_class = CONVERTER_REGISTRY.get(args.conversion_type)
        if not converter_class:
            raise ValueError(f"Unsupported conversion type: {args.conversion_type}")

        # Build kwargs for the converter
        converter_kwargs = {
            'output_dir': args.output_dir,
            'docx_template_path': args.docx_template_path,
            'pptx_template_path': args.pptx_template_path,
            'project_name': args.project_name,
            'author': args.author,
            'email': args.email,
            'mobilephone': args.mobilephone,
            'promote_headings': args.promote_headings,
            'code_highlight_theme': args.code_highlight_theme,
            'poppler_path': args.poppler_path,
            'tesseract_cmd': args.tesseract_cmd,
            # SVG conversion parameters
            'svg_dpi': args.svg_dpi,
            'svg_output_width': args.svg_output_width
        }

        # Extract output_format from conversion_type and pass it through.
        if args.conversion_type.startswith('md-to-'):
            output_format = args.conversion_type.split('-')[-1]
            converter_kwargs['output_format'] = output_format

        # Instantiate the converter
        converter = converter_class(**converter_kwargs)

        # Preparation done
        report_progress("analyzing", 25)

        # Run the conversion
        report_progress("converting", 50)
        output_files = converter.convert(args.input_path)
        success = len(output_files) > 0

        # Report completion
        report_progress("complete", 100)

        # Build the final result
        result = {
            "type": "result",
            "success": success,
            "outputFiles": output_files
        }
        # Base64-encode
        json_str = json.dumps(result, ensure_ascii=False)
        encoded_str = base64.b64encode(json_str.encode('utf-8')).decode('ascii')
        print(encoded_str, flush=True)

    except Exception as e:
        # Report the error
        error_result = {
            "type": "result",
            "success": False,
            "error": str(e)
        }
        # Base64-encode
        json_str = json.dumps(error_result, ensure_ascii=False)
        encoded_str = base64.b64encode(json_str.encode('utf-8')).decode('ascii')
        print(encoded_str, flush=True)
        sys.exit(1)

if __name__ == '__main__':
    main()