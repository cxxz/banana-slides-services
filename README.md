# PDF to Editable PPTX Converter

Convert PDF slides into editable PowerPoint presentations with AI-powered background cleaning.

> **Based on [Banana Slides](https://github.com/Anionex/banana-slides) by Anionex** - An AI-powered presentation generation tool.

## Features

- **PDF to Images**: Converts PDF pages to high-quality PNG images (2x scaling)
- **Text Extraction**: Uses MinerU service to extract text positions and content
- **Background Cleaning**: AI-powered inpainting removes text/icons to create clean backgrounds
- **Editable Output**: Generates PPTX with positioned text boxes over clean backgrounds
- **Template Support**: Apply custom fonts and colors from template PPTX files

## How It Works

1. **PDF → Images**: Convert each PDF page to a PNG image
2. **Parse with MinerU**: Extract text positions, fonts, and content
3. **Inpainting**: Remove text/icons from images using VolcEngine AI
4. **Build PPTX**: Create editable slides with text boxes positioned over clean backgrounds

## Requirements

- Python 3.9+
- [MinerU](https://mineru.net) account for PDF parsing
- (Optional) VolcEngine account for AI inpainting

## Installation

```bash
# Clone the repository
git clone https://github.com/yourusername/pdf-to-pptx.git
cd pdf-to-pptx

# Create virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

## Configuration

Copy `.env.example` to `.env` and configure:

```bash
cp .env.example .env
```

Required:
- `MINERU_TOKEN` - Your MinerU API token

Optional (for AI inpainting):
- `VOLCENGINE_ACCESS_KEY` - VolcEngine access key
- `VOLCENGINE_SECRET_KEY` - VolcEngine secret key

## Usage

Basic usage:
```bash
python pdf_to_pptx.py -i input.pdf
```

With custom output directory:
```bash
python pdf_to_pptx.py -i input.pdf -o ./my_output
```

With template styling:
```bash
python pdf_to_pptx.py -i input.pdf -t ./templates
```

### Options

| Option | Description |
|--------|-------------|
| `-i, --input-file` | Path to input PDF file (required) |
| `-o, --output-dir` | Output directory (default: `./output_files`) |
| `-t, --template-dir` | Template directory with `title-slide.pptx` and `non-title-slide.pptx` |

### Output

- `output_files/<filename>.pptx` - Editable PowerPoint file
- `output_files/images/` - Extracted page images

## Template Support

Place template files in the `templates/` directory:
- `title-slide.pptx` - Template for title slides
- `non-title-slide.pptx` - Template for content slides

The converter will extract fonts and colors from these templates and apply them to the generated presentation.

## Project Structure

```
pdf-to-pptx/
├── pdf_to_pptx.py              # Main CLI entry point
├── config.py                   # Configuration management
├── prompts.py                  # AI prompts
├── export_service.py           # PPTX generation
├── export_service_inpainting.py # Inpainting integration
├── file_parser_service.py      # MinerU PDF parsing
├── image_editability_service.py # Image processing
├── inpainting_service.py       # VolcEngine inpainting
├── ai_providers/               # AI provider implementations
├── utils/                      # Utility functions
└── templates/                  # Template PPTX files
```

## Credits

Most the codebase is adapted from [Banana Slides](https://github.com/Anionex/banana-slides) by [Anionex](https://github.com/Anionex).

## License

This work is licensed under the [Creative Commons Attribution-NonCommercial-ShareAlike 4.0 International License (CC BY-NC-SA 4.0)](https://creativecommons.org/licenses/by-nc-sa/4.0/).

You are free to:
- **Share** — copy and redistribute the material in any medium or format
- **Adapt** — remix, transform, and build upon the material

Under the following terms:
- **Attribution** — You must give appropriate credit and provide a link to the license
- **NonCommercial** — You may not use the material for commercial purposes
- **ShareAlike** — If you remix, transform, or build upon the material, you must distribute your contributions under the same license
