"""
Template Style Extractor - Extract font and color styles from PPTX template files

Uses the existing templates/inventory.py to parse template files and extract
font information from placeholders.
"""
import logging
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional, Tuple, List

logger = logging.getLogger(__name__)


@dataclass
class FontStyle:
    """Font style configuration for a text category (title or body)"""
    name: str                                      # Font family name (e.g., "HPE Graphik")
    fallback_name: str = "Arial"                   # Fallback font if primary not available
    bold: bool = False                             # Whether text should be bold
    italic: bool = False                           # Whether text should be italic
    color_rgb: Optional[Tuple[int, int, int]] = None  # RGB color tuple (r, g, b)


@dataclass
class StyleConfig:
    """Complete style configuration extracted from templates"""
    title_font: FontStyle = field(default_factory=lambda: FontStyle(
        name="Arial",
        bold=True
    ))
    body_font: FontStyle = field(default_factory=lambda: FontStyle(
        name="Arial",
        bold=False
    ))

    # Source template info (for debugging)
    source_template: Optional[str] = None


class TemplateStyleExtractor:
    """
    Extract font and color styles from PPTX template files.

    Uses templates/inventory.py to parse the template and extract font information
    from placeholder shapes (TITLE, BODY, etc.).

    Usage:
        extractor = TemplateStyleExtractor()
        style_config = extractor.extract_styles('/path/to/templates')
    """

    def __init__(self):
        self.logger = logging.getLogger(__name__)

    def extract_styles(
        self,
        template_dir: str,
        content_template: str = "non-title-slide.pptx"
    ) -> StyleConfig:
        """
        Extract styles from template files.

        Args:
            template_dir: Directory containing template PPTX files
            content_template: Filename for content slide template (default: non-title-slide.pptx)

        Returns:
            StyleConfig with extracted font and color information
        """
        template_path = Path(template_dir) / content_template

        if not template_path.exists():
            self.logger.warning(f"Template not found: {template_path}, using HPE defaults")
            return self.get_hpe_default_style()

        try:
            # Import inventory module from templates directory
            templates_dir = Path(__file__).parent.parent / 'templates'
            if str(templates_dir) not in sys.path:
                sys.path.insert(0, str(templates_dir))

            from inventory import extract_text_inventory

            # Extract text inventory from template
            inventory = extract_text_inventory(template_path)

            # Initialize with defaults
            title_font_name = None
            title_font_bold = None
            title_font_color = None
            body_font_name = None
            body_font_bold = None
            body_font_color = None

            # Scan placeholders for font information
            for slide_key, slide_data in inventory.items():
                for shape_key, shape_data in slide_data.items():
                    placeholder_type = shape_data.placeholder_type
                    paragraphs = shape_data.paragraphs

                    if not paragraphs:
                        continue

                    # Get first paragraph's font info
                    first_para = paragraphs[0]
                    font_name = first_para.font_name
                    font_bold = first_para.bold
                    font_color = self._parse_color(first_para.color)

                    if not font_name:
                        continue

                    # Map placeholder type to font category
                    if placeholder_type in ['TITLE', 'CENTER_TITLE']:
                        if not title_font_name:  # Take first match
                            title_font_name = font_name
                            title_font_bold = font_bold
                            title_font_color = font_color
                            self.logger.debug(f"Found title font: {font_name} from {placeholder_type}")

                    elif placeholder_type in ['BODY', 'SUBTITLE', 'FOOTER']:
                        if not body_font_name:  # Take first match
                            body_font_name = font_name
                            body_font_bold = font_bold
                            body_font_color = font_color
                            self.logger.debug(f"Found body font: {font_name} from {placeholder_type}")

            # Build StyleConfig
            config = StyleConfig(source_template=str(template_path))

            if title_font_name:
                config.title_font = FontStyle(
                    name=title_font_name,
                    bold=title_font_bold if title_font_bold is not None else True,
                    color_rgb=title_font_color
                )
            else:
                # Use HPE default for title
                config.title_font = FontStyle(
                    name="HPE Graphik Semibold",
                    bold=True
                )

            if body_font_name:
                config.body_font = FontStyle(
                    name=body_font_name,
                    bold=body_font_bold if body_font_bold is not None else False,
                    color_rgb=body_font_color
                )
            else:
                # Use HPE default for body
                config.body_font = FontStyle(
                    name="HPE Graphik",
                    bold=False
                )

            self.logger.info(f"Extracted styles from {template_path}")
            self.logger.info(f"  Title font: {config.title_font.name} (bold={config.title_font.bold})")
            self.logger.info(f"  Body font: {config.body_font.name} (bold={config.body_font.bold})")

            return config

        except ImportError as e:
            self.logger.warning(f"Could not import inventory module: {e}, using HPE defaults")
            return self.get_hpe_default_style()
        except Exception as e:
            self.logger.warning(f"Failed to extract styles from {template_path}: {e}, using HPE defaults")
            return self.get_hpe_default_style()

    def _parse_color(self, color_str: Optional[str]) -> Optional[Tuple[int, int, int]]:
        """
        Parse color string from inventory to RGB tuple.

        Args:
            color_str: Color string (e.g., "FF0000" for red)

        Returns:
            RGB tuple (r, g, b) or None
        """
        if not color_str:
            return None

        try:
            # Remove any leading # if present
            color_str = color_str.lstrip('#')

            # Parse as hex
            if len(color_str) == 6:
                r = int(color_str[0:2], 16)
                g = int(color_str[2:4], 16)
                b = int(color_str[4:6], 16)
                return (r, g, b)
        except (ValueError, IndexError):
            pass

        return None

    @staticmethod
    def get_hpe_default_style() -> StyleConfig:
        """
        Return HPE brand default style configuration.

        Use this when templates cannot be parsed or as fallback.
        """
        return StyleConfig(
            title_font=FontStyle(
                name="HPE Graphik Semibold",
                fallback_name="Arial",
                bold=True,
                color_rgb=None  # Use default black
            ),
            body_font=FontStyle(
                name="HPE Graphik",
                fallback_name="Arial",
                bold=False,
                color_rgb=None  # Use default black
            ),
            source_template="HPE defaults"
        )
