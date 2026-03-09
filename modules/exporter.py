"""
PowerPoint slide export logic.
Exports slides from a PowerPoint file to PNG images (Page 1, Page 2, etc.)
"""

# Standard library for path and filesystem operations
import os


def export_powerpoint_slides(
    pptx_path: str,
    output_dir: str,
    progress_callback=None,
) -> tuple[int, str | None]:
    """
    Export all slides from a PowerPoint file to images.

    Args:
        pptx_path: Path to the PowerPoint file (.ppt or .pptx)
        output_dir: Directory to save the exported images
        progress_callback: Optional callback(current, total) called for each slide

    Returns:
        Tuple of (slide_count, error_message). error_message is None on success.
    """
    # Lazy import: pywin32 is Windows-only; fail gracefully if not installed
    try:
        import win32com.client
    except ImportError:
        return 0, "pywin32 is not installed. Run: pip install pywin32"

    # Normalize paths to absolute for consistent behavior
    pptx_path = os.path.abspath(pptx_path)
    output_dir = os.path.abspath(output_dir)

    # Validate that the source file exists before starting
    if not os.path.exists(pptx_path):
        return 0, f"File not found: {pptx_path}"

    # Create output directory (and any parent dirs) if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)

    # Initialize references for cleanup in finally block
    powerpoint = None
    presentation = None

    try:
        # Launch PowerPoint via COM automation (requires PowerPoint installed)
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        # Visible=1 avoids COM errors on Python 3; PowerPoint window may flash briefly
        powerpoint.Visible = 1

        # Open the presentation without showing its window
        presentation = powerpoint.Presentations.Open(pptx_path, WithWindow=False)
        # PowerPoint Slides collection is 1-indexed
        slide_count = presentation.Slides.Count

        # Loop through each slide (1 to slide_count inclusive)
        for i in range(1, slide_count + 1):
            # Get slide by index (1-based)
            slide = presentation.Slides(i)
            # Build output path: Page 1.png, Page 2.png, etc.
            output_path = os.path.join(output_dir, f"Page {i}.png")
            # Export slide to PNG using PowerPoint's built-in export
            slide.Export(output_path, "PNG")
            # Notify caller of progress (e.g., for GUI updates)
            if progress_callback:
                progress_callback(i, slide_count)

        # Success: return count and no error
        return slide_count, None

    except Exception as e:
        # Capture any error and return it as a string
        return 0, str(e)

    finally:
        # Always clean up: close presentation and quit PowerPoint
        try:
            if presentation:
                presentation.Close()
            if powerpoint:
                powerpoint.Quit()
        except Exception:
            # Ignore cleanup errors (e.g., if already closed)
            pass
