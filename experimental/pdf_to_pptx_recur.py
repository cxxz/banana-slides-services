#!/usr/bin/env python3
"""
PDF to Editable PPTX Converter (Recursive Analysis)

Converts a PDF file containing slides into a high-quality editable PPTX file by:
1. Converting PDF pages to images
2. Using ImageEditabilityService for recursive element extraction
3. Generating clean backgrounds via inpainting
4. Optionally using Baidu OCR for editable table cells
5. Creating an editable PPTX with positioned elements

This script uses the advanced recursive analysis pipeline for better results
compared to the simpler pdf_to_pptx.py script.

Usage:
    conda activate web
    python pdf_to_pptx_recur.py input_files/test_slides.pdf

Output will be saved to output_files/<input_name>.pptx

Environment Variables (loaded from .env):
    MINERU_TOKEN           - Required: MinerU API token
    MINERU_API_BASE        - Optional: MinerU API base (default: https://mineru.net)
    VOLCENGINE_ACCESS_KEY  - Required: VolcEngine inpainting access key
    VOLCENGINE_SECRET_KEY  - Required: VolcEngine inpainting secret key
    BAIDU_OCR_API_KEY      - Optional: Baidu OCR API key (for editable tables)
    BAIDU_OCR_API_SECRET   - Optional: Baidu OCR API secret (for BCEv3 signing)
"""

import os
import sys
import logging
from pathlib import Path
from typing import List

# Add the current directory to the path for imports
script_dir = Path(__file__).parent
sys.path.insert(0, str(script_dir))

# Load environment variables from .env
from dotenv import load_dotenv
load_dotenv(script_dir / '.env')

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


def pdf_to_images(pdf_path: str, output_dir: str) -> List[str]:
    """
    Convert PDF to images using PyMuPDF

    Args:
        pdf_path: Path to input PDF file
        output_dir: Directory to save output images

    Returns:
        List of paths to generated images
    """
    import fitz  # PyMuPDF

    os.makedirs(output_dir, exist_ok=True)
    doc = fitz.open(pdf_path)
    image_paths = []

    logger.info(f"  PDF has {len(doc)} pages")

    for i, page in enumerate(doc):
        # Render at native PDF size to preserve original page dimensions
        pix = page.get_pixmap()
        img_path = os.path.join(output_dir, f"page_{i:03d}.png")
        pix.save(img_path)
        image_paths.append(img_path)
        logger.info(f"  Saved page {i+1}/{len(doc)}: {img_path}")

    doc.close()
    return image_paths


def check_environment():
    """
    Check that required environment variables are set and log status of optional ones.
    """
    # Required
    mineru_token = os.getenv('MINERU_TOKEN')
    volcengine_ak = os.getenv('VOLCENGINE_ACCESS_KEY')
    volcengine_sk = os.getenv('VOLCENGINE_SECRET_KEY')

    # Optional
    baidu_api_key = os.getenv('BAIDU_OCR_API_KEY')
    baidu_api_secret = os.getenv('BAIDU_OCR_API_SECRET')

    logger.info("Environment check:")
    logger.info(f"  MINERU_TOKEN: {'Set' if mineru_token else 'NOT SET (required)'}")
    logger.info(f"  VOLCENGINE_ACCESS_KEY: {'Set' if volcengine_ak else 'NOT SET (required)'}")
    logger.info(f"  VOLCENGINE_SECRET_KEY: {'Set' if volcengine_sk else 'NOT SET (required)'}")
    logger.info(f"  BAIDU_OCR_API_KEY: {'Set (table OCR enabled)' if baidu_api_key else 'Not set (table OCR disabled)'}")
    logger.info(f"  BAIDU_OCR_API_SECRET: {'Set' if baidu_api_secret else 'Not set'}")

    if not mineru_token:
        raise ValueError("MINERU_TOKEN not configured in .env file")
    if not volcengine_ak or not volcengine_sk:
        raise ValueError("VOLCENGINE credentials not configured in .env file")

    return {
        'mineru_token': mineru_token,
        'mineru_api_base': os.getenv('MINERU_API_BASE', 'https://mineru.net'),
        'baidu_ocr_enabled': bool(baidu_api_key)
    }


def main(pdf_path: str, output_dir: str = None, max_depth: int = 3, max_workers: int = 4):
    """
    Main conversion pipeline using recursive analysis

    Args:
        pdf_path: Path to input PDF file
        output_dir: Output directory (default: ./output_files)
        max_depth: Maximum recursion depth for element extraction (default: 3)
        max_workers: Number of parallel workers for image processing (default: 4)
    """
    from PIL import Image
    from image_editability_service import get_image_editability_service
    from export_service import ExportService

    # Validate input file exists
    if not os.path.exists(pdf_path):
        logger.error(f"Input file not found: {pdf_path}")
        sys.exit(1)

    # Check environment
    env_config = check_environment()

    # Setup output directories
    if output_dir is None:
        output_dir = script_dir / "output_files"
    else:
        output_dir = Path(output_dir)
    images_dir = output_dir / "images"
    os.makedirs(images_dir, exist_ok=True)

    # Derive output filename from input
    pdf_name = Path(pdf_path).stem
    output_pptx = output_dir / f"{pdf_name}.pptx"

    # Determine upload folder (for MinerU results)
    upload_folder = os.getenv('UPLOAD_FOLDER')
    if not upload_folder:
        # Match the path calculation used by file_parser_service.py
        project_root = script_dir.parent.parent
        upload_folder = str(project_root / 'uploads')

    logger.info("=" * 60)
    logger.info("PDF to Editable PPTX Converter (Recursive Analysis)")
    logger.info("=" * 60)
    logger.info(f"Input: {pdf_path}")
    logger.info(f"Output directory: {output_dir}")
    logger.info(f"Max recursion depth: {max_depth}")
    logger.info(f"Max workers: {max_workers}")
    logger.info(f"Baidu OCR: {'Enabled' if env_config['baidu_ocr_enabled'] else 'Disabled'}")
    logger.info("")

    # Step 1: PDF to images
    logger.info("Step 1: Converting PDF to images...")
    image_paths = pdf_to_images(pdf_path, str(images_dir))
    logger.info(f"  Created {len(image_paths)} images")

    # Get dimensions from first image
    with Image.open(image_paths[0]) as img:
        slide_width, slide_height = img.size
    logger.info(f"  Slide dimensions: {slide_width}x{slide_height}")
    logger.info("")

    # Step 2: Initialize ImageEditabilityService
    logger.info("Step 2: Initializing ImageEditabilityService...")
    service = get_image_editability_service(
        mineru_token=env_config['mineru_token'],
        mineru_api_base=env_config['mineru_api_base'],
        max_depth=max_depth,
        upload_folder=upload_folder
    )
    logger.info(f"  Service initialized with max_depth={max_depth}")
    logger.info("")

    # Step 3: Process all images with recursive analysis
    logger.info("Step 3: Processing images with recursive analysis...")
    logger.info("  This includes: MinerU parsing, element extraction, inpainting, and optional Baidu OCR")
    editable_images = service.make_multi_images_editable(
        image_paths=image_paths,
        parallel=True,
        max_workers=max_workers
    )
    logger.info(f"  Processed {len(editable_images)} pages")

    # Log analysis summary
    for i, editable_img in enumerate(editable_images):
        elem_count = len(editable_img.elements)
        has_clean_bg = bool(editable_img.clean_background)
        logger.info(f"  Page {i+1}: {elem_count} elements, clean_background={'Yes' if has_clean_bg else 'No'}")
    logger.info("")

    # Step 4: Create editable PPTX
    logger.info("Step 4: Creating editable PPTX with recursive elements...")
    ExportService.create_editable_pptx_with_recursive_analysis(
        editable_images=editable_images,
        output_file=str(output_pptx),
        slide_width_pixels=slide_width,
        slide_height_pixels=slide_height
    )
    logger.info(f"  PPTX saved to: {output_pptx}")

    # Summary
    logger.info("")
    logger.info("=" * 60)
    logger.info("Conversion complete!")
    logger.info("=" * 60)
    logger.info(f"Output: {output_pptx}")
    logger.info(f"Images kept at: {images_dir}")

    # Count total elements
    total_elements = sum(len(img.elements) for img in editable_images)
    total_with_clean_bg = sum(1 for img in editable_images if img.clean_background)
    logger.info(f"Total elements extracted: {total_elements}")
    logger.info(f"Pages with clean backgrounds: {total_with_clean_bg}/{len(editable_images)}")


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="Convert PDF slides to editable PPTX using recursive analysis"
    )
    parser.add_argument(
        "-i", "--input-file",
        required=True,
        help="Path to input PDF file"
    )
    parser.add_argument(
        "-o", "--output-dir",
        default=None,
        help="Output directory (default: ./output_files)"
    )
    parser.add_argument(
        "--max-depth",
        type=int,
        default=3,
        help="Maximum recursion depth for element extraction (default: 3)"
    )
    parser.add_argument(
        "--max-workers",
        type=int,
        default=4,
        help="Number of parallel workers for image processing (default: 4)"
    )

    args = parser.parse_args()

    try:
        main(
            args.input_file,
            output_dir=args.output_dir,
            max_depth=args.max_depth,
            max_workers=args.max_workers
        )
    except Exception as e:
        logger.error(f"Conversion failed: {e}", exc_info=True)
        sys.exit(1)
