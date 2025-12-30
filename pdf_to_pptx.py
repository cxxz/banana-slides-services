#!/usr/bin/env python3
"""
PDF to Editable PPTX Converter

Converts a PDF file containing slides into an editable PPTX file by:
1. Converting PDF pages to images
2. Parsing the PDF with MinerU to extract text positions
3. Generating clean backgrounds using inpainting (removing text/icons)
4. Creating an editable PPTX with positioned text boxes and template styling

Usage:
    conda activate web
    python pdf_to_pptx.py -i input_files/test_slides.pdf

    # With custom template directory:
    python pdf_to_pptx.py -i input_files/test_slides.pdf -t /path/to/templates

Output will be saved to output_files/<input_name>.pptx

Template Support:
    Place title-slide.pptx and non-title-slide.pptx in the templates/ directory.
    Fonts and colors will be extracted and applied to the generated PPTX.
    Default fonts: HPE Graphik Semibold (titles), HPE Graphik (body)
"""

import os
import sys
import logging
from pathlib import Path
from typing import List, Optional

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
        # 2x scaling for high quality output
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
        img_path = os.path.join(output_dir, f"page_{i:03d}.png")
        pix.save(img_path)
        image_paths.append(img_path)
        logger.info(f"  Saved page {i+1}/{len(doc)}: {img_path}")

    doc.close()
    return image_paths


def parse_pdf_with_mineru(pdf_path: str, filename: str) -> str:
    """
    Parse PDF using MinerU service

    Args:
        pdf_path: Path to input PDF file
        filename: Original filename

    Returns:
        extract_id: ID for the extracted results directory
    """
    from file_parser_service import FileParserService
    from utils.self_hosted_mineru import is_self_hosted_mineru, parse_pdf_via_self_hosted_mineru

    mineru_token = os.getenv('MINERU_TOKEN')
    mineru_api_base = os.getenv('MINERU_API_BASE', 'https://mineru.net')

    # Self-hosted MinerU (/file_parse) does not use the cloud token flow.
    if is_self_hosted_mineru(mineru_api_base):
        upload_folder = os.getenv('UPLOAD_FOLDER')
        if upload_folder:
            upload_root = Path(upload_folder)
        else:
            # Match the default used later in this script (and FileParserService): ~/uploads
            upload_root = Path(__file__).resolve().parent.parent.parent / "uploads"

        extract_id, _extract_dir = parse_pdf_via_self_hosted_mineru(
            Path(pdf_path),
            mineru_api_base=mineru_api_base,
            upload_folder=upload_root,
        )
        logger.info(f"  MinerU (self-hosted) extract_id: {extract_id}")
        return extract_id

    if not mineru_token:
        raise ValueError("MINERU_TOKEN not configured in .env file (required for mineru.net)")

    parser = FileParserService(
        mineru_token=mineru_token,
        mineru_api_base=mineru_api_base
    )

    batch_id, markdown, extract_id, error, failed_count = parser.parse_file(pdf_path, filename)

    if error:
        raise ValueError(f"MinerU parsing failed: {error}")

    if not extract_id:
        raise ValueError("MinerU parsing failed: no extract_id returned")

    logger.info(f"  MinerU batch_id: {batch_id}")
    logger.info(f"  MinerU extract_id: {extract_id}")

    return extract_id


def generate_clean_backgrounds(image_paths: List[str], mineru_result_dir: str) -> List[str]:
    """
    Generate clean backgrounds using inpainting

    Removes text, icons, and overlays from slide images to create clean backgrounds.
    Falls back to original images if inpainting is not configured or fails.

    Args:
        image_paths: List of paths to original page images
        mineru_result_dir: Directory containing MinerU parsing results

    Returns:
        List of paths to clean background images
    """
    from export_service_inpainting import InpaintingExportHelper

    # Check if VolcEngine credentials are configured
    volcengine_access_key = os.getenv('VOLCENGINE_ACCESS_KEY')
    volcengine_secret_key = os.getenv('VOLCENGINE_SECRET_KEY')
    use_inpainting = bool(volcengine_access_key and volcengine_secret_key)

    if not use_inpainting:
        logger.warning("  VolcEngine credentials not configured, using original images as backgrounds")
        return image_paths

    logger.info("  Using VolcEngine inpainting to generate clean backgrounds...")

    return InpaintingExportHelper.generate_clean_backgrounds_with_inpainting(
        image_paths=image_paths,
        mineru_result_dir=mineru_result_dir,
        use_inpainting=use_inpainting
    )


def create_editable_pptx(
    mineru_result_dir: str,
    output_file: str,
    slide_width: int,
    slide_height: int,
    background_images: List[str],
    template_dir: str = None
):
    """
    Create editable PPTX from MinerU results

    Args:
        mineru_result_dir: Directory containing MinerU parsing results
        output_file: Path for output PPTX file
        slide_width: Width of slides in pixels
        slide_height: Height of slides in pixels
        background_images: List of paths to background images
        template_dir: Optional directory containing template PPTX files for styling
    """
    from export_service import ExportService

    # Create output directory if needed
    output_dir = os.path.dirname(output_file)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)

    ExportService.create_editable_pptx_from_mineru(
        mineru_result_dir=mineru_result_dir,
        output_file=output_file,
        slide_width_pixels=slide_width,
        slide_height_pixels=slide_height,
        background_images=background_images,
        template_dir=template_dir
    )


def main(pdf_path: str, output_dir: str = None, template_dir: str = None, mineru_results_dir: str = None):
    """
    Main conversion pipeline

    Args:
        pdf_path: Path to input PDF file
        output_dir: Output directory (default: ./output_files)
        template_dir: Template directory (default: ./templates)
    """
    from PIL import Image

    # Validate input file exists
    if not os.path.exists(pdf_path):
        logger.error(f"Input file not found: {pdf_path}")
        sys.exit(1)

    # Setup output directories
    if output_dir is None:
        output_dir = script_dir / "output_files"
    else:
        output_dir = Path(output_dir)
    images_dir = output_dir / "images"
    os.makedirs(images_dir, exist_ok=True)

    # Setup template directory (default to ./templates)
    if template_dir is None:
        default_template_dir = script_dir / "templates"
        if default_template_dir.exists():
            template_dir = str(default_template_dir)
            logger.info(f"Using default template directory: {template_dir}")
        else:
            logger.info("No template directory found, using default styling")

    # Derive output filename from input
    pdf_name = Path(pdf_path).stem
    output_pptx = output_dir / f"{pdf_name}.pptx"

    logger.info("=" * 60)
    logger.info("PDF to Editable PPTX Converter")
    logger.info("=" * 60)
    logger.info(f"Input: {pdf_path}")
    logger.info(f"Output directory: {output_dir}")
    logger.info("")

    # Step 1: PDF to images (kept in output_files/images/)
    logger.info("Step 1: Converting PDF to images...")
    image_paths = pdf_to_images(pdf_path, str(images_dir))
    logger.info(f"  Created {len(image_paths)} images")

    # Get dimensions from first image
    with Image.open(image_paths[0]) as img:
        slide_width, slide_height = img.size
    logger.info(f"  Slide dimensions: {slide_width}x{slide_height}")
    logger.info("")

    # Step 2: Parse with MinerU (or reuse an existing MinerU result dir)
    if mineru_results_dir:
        mineru_dir = Path(mineru_results_dir).expanduser().resolve()
        logger.info("Step 2: Skipping MinerU parsing (using existing results)...")
        logger.info(f"  MinerU results: {mineru_dir}")

        if not mineru_dir.exists() or not mineru_dir.is_dir():
            raise ValueError(f"--mineru-results-dir does not exist or is not a directory: {mineru_dir}")

        content_list_files = list(mineru_dir.glob("*_content_list.json"))
        if not content_list_files:
            raise ValueError(f"No *_content_list.json found in --mineru-results-dir: {mineru_dir}")

        if not (mineru_dir / "layout.json").exists():
            logger.warning(f"  layout.json not found in {mineru_dir} (will fall back to content_list coords)")

        mineru_result_dir = str(mineru_dir)
        logger.info("")
    else:
        logger.info("Step 2: Parsing PDF with MinerU...")
        extract_id = parse_pdf_with_mineru(pdf_path, os.path.basename(pdf_path))

        # Determine MinerU result directory
        # Note: file_parser_service.py uses a specific path calculation:
        #   current_file.parent.parent.parent / 'uploads' / 'mineru_files'
        # This means: script_dir -> parent -> parent -> uploads
        # For /Users/bill/workspace/banana-slides-services, this becomes /Users/bill/uploads
        upload_folder = os.getenv('UPLOAD_FOLDER')
        if not upload_folder:
            # Match the path calculation in file_parser_service.py
            # It goes: file_parser_service.py -> parent.parent (workspace) -> parent (bill) -> uploads
            project_root = script_dir.parent.parent  # Goes up 2 levels from script_dir
            upload_folder = str(project_root / 'uploads')

        mineru_result_dir = os.path.join(upload_folder, 'mineru_files', extract_id)
        logger.info(f"  MinerU results: {mineru_result_dir}")
        logger.info("")

    # Step 3: Generate clean backgrounds
    logger.info("Step 3: Generating clean backgrounds...")
    clean_bg_paths = generate_clean_backgrounds(image_paths, mineru_result_dir)
    logger.info(f"  Generated {len(clean_bg_paths)} backgrounds")
    logger.info("")

    # Step 4: Create editable PPTX with template styling
    logger.info("Step 4: Creating editable PPTX with template styling...")
    create_editable_pptx(
        mineru_result_dir=mineru_result_dir,
        output_file=str(output_pptx),
        slide_width=slide_width,
        slide_height=slide_height,
        background_images=clean_bg_paths,
        template_dir=template_dir
    )
    logger.info(f"  PPTX saved to: {output_pptx}")

    # Summary
    logger.info("")
    logger.info("=" * 60)
    logger.info("Conversion complete!")
    logger.info("=" * 60)
    logger.info(f"Output: {output_pptx}")
    logger.info(f"Images kept at: {images_dir}")
    logger.info(f"MinerU results at: {mineru_result_dir}")


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="Convert PDF slides to editable PPTX with template styling"
    )
    parser.add_argument(
        "-i", "--input-file",
        required=True,
        help="Path to input PDF file"
    )
    parser.add_argument(
        "--mineru-results-dir",
        default=None,
        help="If provided, skip MinerU parsing and use this existing MinerU result directory "
             "(must contain *_content_list.json).",
    )
    parser.add_argument(
        "-o", "--output-dir",
        default=None,
        help="Output directory (default: ./output_files)"
    )
    parser.add_argument(
        "-t", "--template-dir",
        default=None,
        help="Template directory containing title-slide.pptx and non-title-slide.pptx "
             "(default: ./templates)"
    )

    args = parser.parse_args()

    try:
        main(args.input_file, output_dir=args.output_dir, template_dir=args.template_dir, mineru_results_dir=args.mineru_results_dir)
    except Exception as e:
        logger.error(f"Conversion failed: {e}", exc_info=True)
        sys.exit(1)
