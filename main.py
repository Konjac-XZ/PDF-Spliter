"""PDF Dot Matrix Overlay Tool.

Adds a centered dot matrix overlay to PDF documents.
"""

from __future__ import annotations

import argparse
import copy
import math
import sys
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path

from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen.canvas import Canvas


# Unit conversion constants
POINTS_PER_INCH = 72
MM_PER_INCH = 25.4
MM_TO_POINTS = POINTS_PER_INCH / MM_PER_INCH


def mm_to_points(millimeters: float) -> float:
    """Convert millimeters to PDF points."""
    return millimeters * MM_TO_POINTS


def points_to_mm(points: float) -> float:
    """Convert PDF points to millimeters."""
    return points / MM_TO_POINTS


@dataclass(frozen=True)
class DotMatrixConfig:
    """Configuration for the dot matrix overlay."""

    dot_spacing_mm: float = 10.0  # 1 cm between dots
    dot_diameter_mm: float = 0.4  # 0.4 mm diameter
    dot_color_hex: str = "#dddddd"  # Light gray

    @property
    def dot_radius_mm(self) -> float:
        """Get the dot radius in millimeters."""
        return self.dot_diameter_mm / 2

    @property
    def dot_color_rgb(self) -> tuple[float, float, float]:
        """Convert hex color to RGB tuple (0-1 range for reportlab)."""
        hex_color = self.dot_color_hex.lstrip("#")
        r = int(hex_color[0:2], 16) / 255
        g = int(hex_color[2:4], 16) / 255
        b = int(hex_color[4:6], 16) / 255
        return (r, g, b)


def calculate_dot_positions(dimension_mm: float, spacing_mm: float) -> list[float]:
    """Calculate centered dot positions along a dimension.

    For a given dimension, calculates positions such that:
    1. Dots are spaced exactly `spacing_mm` apart
    2. The dot pattern is centered with equal margins on both sides

    Args:
        dimension_mm: The total dimension in millimeters
        spacing_mm: The spacing between dots in millimeters

    Returns:
        List of dot positions in millimeters from the origin
    """
    num_intervals = math.floor(dimension_mm / spacing_mm)
    span = num_intervals * spacing_mm
    margin = (dimension_mm - span) / 2
    return [margin + i * spacing_mm for i in range(num_intervals + 1)]


def create_half_white_overlay(
    width_pt: float,
    height_pt: float,
    cover_side: str,
) -> BytesIO:
    """Create a PDF overlay with white rectangle covering half the page.

    Args:
        width_pt: Page width in points
        height_pt: Page height in points
        cover_side: Which side to cover with white ("left" or "right")

    Returns:
        BytesIO containing the overlay PDF
    """
    buffer = BytesIO()
    canvas = Canvas(buffer, pagesize=(width_pt, height_pt))
    canvas.setFillColorRGB(1, 1, 1)  # White

    half_width = width_pt / 2
    if cover_side == "left":
        canvas.rect(0, 0, half_width, height_pt, stroke=0, fill=1)
    else:  # cover_side == "right"
        canvas.rect(half_width, 0, half_width, height_pt, stroke=0, fill=1)

    canvas.save()
    buffer.seek(0)
    return buffer


def create_dot_matrix_overlay(
    width_pt: float,
    height_pt: float,
    config: DotMatrixConfig,
    split_border: bool = False,
) -> BytesIO:
    """Create a PDF overlay with a centered dot matrix.

    Args:
        width_pt: Page width in points
        height_pt: Page height in points
        config: Dot matrix configuration
        split_border: If True, draw a vertical line at center connecting first and last row

    Returns:
        BytesIO containing the overlay PDF
    """
    buffer = BytesIO()
    canvas = Canvas(buffer, pagesize=(width_pt, height_pt))

    # Convert dimensions to mm for calculation
    width_mm = points_to_mm(width_pt)
    height_mm = points_to_mm(height_pt)

    # Calculate dot positions
    x_positions = calculate_dot_positions(width_mm, config.dot_spacing_mm)
    y_positions = calculate_dot_positions(height_mm, config.dot_spacing_mm)

    # Set dot appearance
    r, g, b = config.dot_color_rgb
    canvas.setFillColorRGB(r, g, b)
    canvas.setStrokeColorRGB(r, g, b)

    radius_pt = mm_to_points(config.dot_radius_mm)

    # Draw dots at each intersection
    for x_mm in x_positions:
        for y_mm in y_positions:
            x_pt = mm_to_points(x_mm)
            y_pt = mm_to_points(y_mm)
            canvas.circle(x_pt, y_pt, radius_pt, stroke=0, fill=1)

    # Draw split border line if enabled
    if split_border and y_positions:
        center_x_pt = width_pt / 2
        first_y_pt = mm_to_points(y_positions[0])
        last_y_pt = mm_to_points(y_positions[-1])

        # Line width matches dot diameter
        canvas.setLineWidth(mm_to_points(config.dot_diameter_mm))
        canvas.setStrokeColorRGB(r, g, b)
        canvas.line(center_x_pt, first_y_pt, center_x_pt, last_y_pt)

    canvas.save()
    buffer.seek(0)
    return buffer


def process_pdf(
    input_path: Path,
    output_path: Path,
    config: DotMatrixConfig,
    split: bool = False,
    split_border: bool = False,
) -> None:
    """Process a PDF file and add dot matrix overlay to each page.

    Args:
        input_path: Path to the input PDF file
        output_path: Path for the output PDF file
        config: Dot matrix configuration
        split: If True, split each page horizontally into two output pages
        split_border: If True (and split is True), draw a border line at center
    """
    reader = PdfReader(input_path)
    writer = PdfWriter()

    for page in reader.pages:
        # Get page dimensions from MediaBox
        media_box = page.mediabox
        width_pt = float(media_box.width)
        height_pt = float(media_box.height)

        # Create dot matrix overlay for this page's dimensions
        # Only add split border when both split and split_border are enabled
        dot_overlay_buffer = create_dot_matrix_overlay(
            width_pt, height_pt, config, split_border=(split and split_border)
        )

        if split:
            # Create two output pages per input page
            for cover_side in ("right", "left"):
                # Start with white overlay as base, merge original page UNDER it
                white_buffer = create_half_white_overlay(width_pt, height_pt, cover_side)
                white_reader = PdfReader(white_buffer)
                output_page = white_reader.pages[0]

                # Merge original page content UNDER the white overlay
                output_page.merge_page(page, over=False)

                # Apply dot matrix overlay on top
                dot_overlay_buffer.seek(0)
                dot_reader = PdfReader(dot_overlay_buffer)
                output_page.merge_page(dot_reader.pages[0])

                writer.add_page(output_page)
        else:
            # Normal mode: just add dot matrix overlay
            dot_reader = PdfReader(dot_overlay_buffer)
            page.merge_page(dot_reader.pages[0])
            writer.add_page(page)

    with open(output_path, "wb") as output_file:
        writer.write(output_file)


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    """Parse command line arguments."""
    parser = argparse.ArgumentParser(
        description="Add a centered dot matrix overlay to a PDF document.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument(
        "input",
        type=Path,
        help="Input PDF file path",
    )
    parser.add_argument(
        "output",
        type=Path,
        help="Output PDF file path",
    )
    parser.add_argument(
        "--spacing",
        type=float,
        default=DotMatrixConfig.dot_spacing_mm,
        metavar="MM",
        help="Distance between dots in millimeters",
    )
    parser.add_argument(
        "--diameter",
        type=float,
        default=DotMatrixConfig.dot_diameter_mm,
        metavar="MM",
        help="Dot diameter in millimeters",
    )
    parser.add_argument(
        "--color",
        type=str,
        default=DotMatrixConfig.dot_color_hex,
        metavar="HEX",
        help="Dot color in hex format (e.g., #dddddd)",
    )
    parser.add_argument(
        "--split",
        action="store_true",
        help="Split each page horizontally, outputting left and right halves as separate pages",
    )
    parser.add_argument(
        "--split-border",
        action="store_true",
        help="Draw a vertical border line at center (requires --split)",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    """Main entry point."""
    args = parse_args(argv)

    if not args.input.exists():
        print(f"Error: Input file not found: {args.input}", file=sys.stderr)
        return 1

    config = DotMatrixConfig(
        dot_spacing_mm=args.spacing,
        dot_diameter_mm=args.diameter,
        dot_color_hex=args.color,
    )

    try:
        process_pdf(args.input, args.output, config, split=args.split, split_border=args.split_border)
        print(f"Successfully created: {args.output}")
        return 0
    except Exception as e:
        print(f"Error processing PDF: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
