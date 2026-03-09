# PowerPoint Slide Exporter — Architecture & Flow

## Mindmap: How the Program Works

```mermaid
mindmap
  root((PowerPoint Slide Exporter))
    Entry Point
      export_slides.py
      Imports SlideExporterApp
      Creates app instance
      Runs mainloop
    GUI Layer
      SlideExporterApp
        File Selection
          Add file(s) button
          Remove selected
          Clear all
          Listbox + scrollbar
        Output
          Folder path entry
          Browse button
        Progress
          Status label
          Progress bar
        Export
          Validate inputs
          Call exporter per file
          Update progress
          Show dialogs
    Export Logic
      export_powerpoint_slides
        Check pywin32
        Validate paths
        Launch PowerPoint COM
        Open presentation
        Loop slides 1 to N
          Export to Page N.png
          Call progress callback
        Cleanup
          Close presentation
          Quit PowerPoint
    Output
      Subfolder per file
      Page 1.png, Page 2.png...
```

---

## Modularization Recommendation

**Current structure is appropriate for this scale.** Further splitting would add files and indirection without clear benefit:

- **exporter.py** (~80 lines) — Single responsibility, easy to test
- **gui.py** (~150 lines) — One cohesive UI class
- **export_slides.py** — Minimal entry point

Splitting further (e.g., separate `file_selector.py`, `progress.py`) would create many small modules with tight coupling. Consider more modularization only if the app grows (e.g., multiple export formats, batch presets, or a plugin system).

---

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

---

## Module Dependency Graph

```mermaid
flowchart LR
    subgraph Modules
        MAIN[export_slides.py]
        GUI[gui.py]
        EXP[exporter.py]
    end

    MAIN -->|imports| GUI
    GUI -->|imports| EXP
    EXP -.->|uses| PPT[PowerPoint COM]
```

---

## Data Flow (Export Sequence)

```mermaid
sequenceDiagram
    participant User
    participant GUI
    participant Exporter
    participant PowerPoint

    User->>GUI: Click "Export Slides to Images"
    GUI->>GUI: Validate files & output folder
    GUI->>GUI: Disable button, reset progress

    loop For each .ppt/.pptx file
        GUI->>Exporter: export_powerpoint_slides(path, output_dir, callback)
        Exporter->>PowerPoint: Dispatch Application
        Exporter->>PowerPoint: Open presentation

        loop For each slide
            Exporter->>PowerPoint: Export slide to PNG
            Exporter->>GUI: progress_callback(current, total)
            GUI->>User: Update progress bar
        end

        Exporter->>PowerPoint: Close, Quit
        Exporter->>GUI: Return (count, error)
    end

    GUI->>User: Show success/error dialog
```

---

## File Structure

```
export-powerpoint-slides-to-images/
├── export_slides.py    # Entry point: creates app, runs mainloop
├── gui.py              # SlideExporterApp: UI, file selection, export orchestration
├── exporter.py         # export_powerpoint_slides(): COM-based slide export
├── requirements.txt
├── README.md
└── ARCHITECTURE.md     # This file
```

---

## Component Responsibilities

| Component | Responsibility |
|-----------|----------------|
| **export_slides.py** | Bootstrap: import GUI, instantiate, run |
| **gui.py** | File selection, output folder, progress display, validation, calling exporter |
| **exporter.py** | PowerPoint COM automation, slide export to PNG, error handling |
