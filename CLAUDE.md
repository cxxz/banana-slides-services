# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a CLI tool that converts PDF slides into editable PowerPoint presentations. It extracts text positions using MinerU, generates clean backgrounds via AI inpainting, and creates PPTX files with positioned text boxes.

Based on [Banana Slides](https://github.com/Anionex/banana-slides) by Anionex.

## Usage

```bash
# Basic usage
python pdf_to_pptx.py -i input.pdf

# With options
python pdf_to_pptx.py -i input.pdf -o ./output -t ./templates
```

## Architecture

### Entry Point
- **pdf_to_pptx.py** - Main CLI script that orchestrates the conversion pipeline

### Core Services
- **ExportService** (`export_service.py`) - Creates PPTX presentations from MinerU results
- **ExportServiceInpainting** (`export_service_inpainting.py`) - Generates clean backgrounds using inpainting
- **FileParserService** (`file_parser_service.py`) - Integrates with MinerU API for PDF parsing
- **InpaintingService** (`inpainting_service.py`) - Image inpainting via VolcEngine API
- **ImageEditabilityService** (`image_editability_service.py`) - Image processing utilities

### Supporting Modules
- **config.py** - Configuration management (environment variables)
- **prompts.py** - AI prompts for background generation

### Utilities (`utils/`)
- **pptx_builder.py** - PowerPoint slide construction
- **template_style_extractor.py** - Extract fonts/colors from template PPTX
- **coordinate_utils.py** - Bounding box and coordinate handling
- **mask_utils.py** - Mask generation for inpainting
- **path_utils.py** - MinerU file path utilities

### AI Providers (`ai_providers/`)
- **image/volcengine_inpainting_provider.py** - VolcEngine inpainting API client

## Key Environment Variables

```bash
# Required
MINERU_TOKEN          # MinerU API token for PDF parsing

# Optional (for AI inpainting)
VOLCENGINE_ACCESS_KEY # VolcEngine access key
VOLCENGINE_SECRET_KEY # VolcEngine secret key

# Optional
MINERU_API_BASE       # MinerU API base URL (default: https://mineru.net)
UPLOAD_FOLDER         # Custom upload folder path
```

## Conversion Pipeline

1. **PDF → Images**: Convert PDF pages to PNG using PyMuPDF (2x scaling)
2. **Parse with MinerU**: Extract text positions, fonts, and content via MinerU API
3. **Generate Backgrounds**: Use VolcEngine inpainting to remove text/icons (optional)
4. **Build PPTX**: Create editable slides with positioned text boxes over backgrounds

## Code Patterns

### Service Instantiation
Services are typically instantiated with configuration from environment:
```python
from file_parser_service import FileParserService

parser = FileParserService(
    mineru_token=os.getenv('MINERU_TOKEN'),
    mineru_api_base=os.getenv('MINERU_API_BASE', 'https://mineru.net')
)
```

### Static Method Usage
Many services use static methods for stateless operations:
```python
from export_service import ExportService

ExportService.create_editable_pptx_from_mineru(
    mineru_result_dir=result_dir,
    output_file=output_path,
    slide_width_pixels=width,
    slide_height_pixels=height,
    background_images=bg_images
)
```

### Retry Logic
Uses `tenacity` for retry logic on external API calls (MinerU, VolcEngine).

## Template Support

Place template files in `templates/` directory:
- `title-slide.pptx` - Template for title slides
- `non-title-slide.pptx` - Template for content slides

Fonts and colors are extracted and applied to generated presentations.

