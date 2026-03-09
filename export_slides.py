"""
PowerPoint Slide to Image Exporter
A GUI application that exports PowerPoint slides to images (Page 1, Page 2, etc.)
Requires Microsoft PowerPoint installed on Windows.
"""

# Import the GUI application class
from modules.gui import SlideExporterApp

# Only run when executed directly (not when imported as a module)
if __name__ == "__main__":
    # Create the application instance
    app = SlideExporterApp()
    # Start the GUI event loop
    app.run()
