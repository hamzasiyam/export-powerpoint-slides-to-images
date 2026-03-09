# PowerPoint Slide to Image Exporter

A simple Python GUI application that exports PowerPoint slides to PNG images. Each slide is saved as **Page 1.png**, **Page 2.png**, and so on.

## Requirements

- **Windows** with **Microsoft PowerPoint** installed
- Python 3.8+

## Installation

```bash
pip install -r requirements.txt
```

## Usage

Run the application:

```bash
python export_slides.py
```

1. Click **Add file(s)...** to select one or more PowerPoint files (.ppt or .pptx)
2. Choose an output folder (or use the default)
3. Click **Export Slides to Images**

Each presentation gets its own subfolder in the output directory (named after the file), containing `Page 1.png`, `Page 2.png`, etc.

## Program Flow Diagram

```mermaid
flowchart TB
    subgraph Entry["export_slides.py (Entry Point)"]
        A[Start] --> B[Import SlideExporterApp]
        B --> C[Create app instance]
        C --> D[app.run]
        D --> E[mainloop]
    end

    subgraph GUI["gui.py (SlideExporterApp)"]
        E --> F[User interacts with UI]
        F --> G{User action?}
        G -->|Add files| H[_add_files]
        G -->|Remove selected| I[_remove_selected]
        G -->|Clear all| J[_clear_files]
        G -->|Browse output| K[_browse_output]
        G -->|Export| L[_export]

        H --> M[Update file listbox]
        I --> M
        J --> M
        K --> N[Update output_dir]

        L --> O{Validation}
        O -->|No files| P[Show warning]
        O -->|No output folder| P
        O -->|OK| Q[For each PowerPoint file]
    end

    subgraph Exporter["exporter.py"]
        Q --> R[export_powerpoint_slides]
        R --> S[Launch PowerPoint via COM]
        S --> T[Open presentation]
        T --> U[Loop: slide 1 to N]
        U --> V[Export slide to Page N.png]
        V --> W[progress_callback]
        W --> U
        U --> X[Close & Quit PowerPoint]
        X --> Y[Return count, error]
    end

    Y --> Z[Update progress bar]
    Z --> AA[Show success/error dialog]
```
