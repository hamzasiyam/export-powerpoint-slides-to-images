"""
GUI for the PowerPoint Slide to Image Exporter.
"""

# Standard library
import os
# Tkinter: built-in GUI framework
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path

# Import the export function from our exporter module
from exporter import export_powerpoint_slides


class SlideExporterApp:
    """Main application window for exporting PowerPoint slides to images."""

    def __init__(self):
        # Create the main application window
        self.root = tk.Tk()
        # Set window title shown in title bar
        self.root.title("PowerPoint Slide to Image Exporter")
        # Set initial size (width x height)
        self.root.geometry("560x320")
        # Allow user to resize the window
        self.root.resizable(True, True)

        # List of full paths to selected PowerPoint files
        self.pptx_files: list[str] = []
        # Tk variable bound to output folder path; default to user's home
        self.output_dir = tk.StringVar(value=os.path.expanduser("~"))

        # Build all UI widgets
        self._build_ui()

    def _build_ui(self):
        # Main container with padding around edges
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- File selection section ---
        # Label above the file list
        ttk.Label(main_frame, text="PowerPoint files:").pack(anchor=tk.W)
        # Frame to hold listbox and scrollbar
        file_frame = ttk.Frame(main_frame)
        file_frame.pack(fill=tk.X, pady=(4, 12))

        # Listbox: shows selected files; EXTENDED allows multi-select
        self.file_listbox = tk.Listbox(file_frame, height=5, selectmode=tk.EXTENDED)
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        # Scrollbar for when many files are added
        scrollbar = ttk.Scrollbar(file_frame, orient=tk.VERTICAL, command=self.file_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        # Link scrollbar to listbox's vertical scroll
        self.file_listbox.config(yscrollcommand=scrollbar.set)

        # Frame for file management buttons
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 12))

        # Add files: opens file picker
        ttk.Button(btn_frame, text="Add file(s)...", command=self._add_files).pack(side=tk.LEFT, padx=(0, 8))
        # Remove selected items from list
        ttk.Button(btn_frame, text="Remove selected", command=self._remove_selected).pack(side=tk.LEFT, padx=(0, 8))
        # Clear entire list
        ttk.Button(btn_frame, text="Clear all", command=self._clear_files).pack(side=tk.LEFT)

        # --- Output directory section ---
        ttk.Label(main_frame, text="Output folder:").pack(anchor=tk.W)
        # Spacer for layout
        ttk.Frame(main_frame).pack(fill=tk.X, pady=(4, 12))

        # Text entry showing/editing output path
        ttk.Entry(main_frame, textvariable=self.output_dir, width=50).pack(fill=tk.X, pady=(0, 4))
        # Browse button: opens folder picker
        ttk.Button(main_frame, text="Browse...", command=self._browse_output).pack(anchor=tk.W)

        # --- Progress section ---
        # StringVar: updates label text when we change it
        self.progress_var = tk.StringVar(value="")
        # Label that shows current status (e.g., "Exporting slide 3/10")
        ttk.Label(main_frame, textvariable=self.progress_var).pack(anchor=tk.W, pady=(12, 4))

        # Progress bar: determinate = we know total and can show percentage
        self.progress_bar = ttk.Progressbar(main_frame, mode="determinate")
        self.progress_bar.pack(fill=tk.X, pady=(0, 12))

        # --- Export button ---
        self.export_btn = ttk.Button(main_frame, text="Export Slides to Images", command=self._export)
        self.export_btn.pack(pady=(8, 0))

    def _add_files(self):
        # Open native file dialog; returns tuple of selected paths
        files = filedialog.askopenfilenames(
            title="Select PowerPoint file(s)",
            filetypes=[
                ("PowerPoint files", "*.ppt *.pptx"),
                ("All files", "*.*"),
            ],
        )
        # Add each file if not already in list
        for f in files:
            if f and f not in self.pptx_files:
                self.pptx_files.append(f)
                # Show only filename in listbox for cleaner display
                self.file_listbox.insert(tk.END, os.path.basename(f))

    def _remove_selected(self):
        # curselection() returns indices of selected items (e.g., [0, 2])
        selected = list(self.file_listbox.curselection())
        # Reverse order so deleting doesn't shift indices
        for i in reversed(selected):
            self.file_listbox.delete(i)
            del self.pptx_files[i]

    def _clear_files(self):
        # Remove all items from listbox (0 to END)
        self.file_listbox.delete(0, tk.END)
        # Clear the internal list
        self.pptx_files.clear()

    def _browse_output(self):
        # Open native folder picker
        folder = filedialog.askdirectory(initialdir=self.output_dir.get())
        if folder:
            # Update the StringVar (and thus the Entry) with chosen path
            self.output_dir.set(folder)

    def _export(self):
        # Validate: must have at least one file
        if not self.pptx_files:
            messagebox.showwarning("No files", "Please add at least one PowerPoint file.")
            return

        # Get output path and validate it's not empty
        output_base = self.output_dir.get().strip()
        if not output_base:
            messagebox.showwarning("No output folder", "Please select an output folder.")
            return

        # Disable button during export to prevent double-clicks
        self.export_btn.config(state=tk.DISABLED)
        # Show initial status
        self.progress_var.set("Exporting...")
        self.progress_bar["value"] = 0
        # Force UI update before long operation
        self.root.update()

        total_slides = 0
        errors = []

        # Process each PowerPoint file
        for pptx_path in self.pptx_files:
            # Use filename (without extension) as subfolder name
            base_name = Path(pptx_path).stem
            # Each file gets its own subfolder: output_base/presentation_name/
            output_dir = os.path.join(output_base, base_name)

            # Callback to update progress bar and status during export
            def update_progress(current, total):
                # Calculate percentage (0-100)
                pct = (current / total) * 100 if total else 0
                # Update status text
                self.progress_var.set(f"Exporting {base_name}: slide {current}/{total}")
                # Update progress bar
                self.progress_bar["value"] = pct
                # Refresh UI so user sees updates
                self.root.update()

            # Call exporter; receives (count, error)
            count, err = export_powerpoint_slides(pptx_path, output_dir, update_progress)

            if err:
                # Collect errors to show at end
                errors.append(f"{os.path.basename(pptx_path)}: {err}")
            else:
                total_slides += count

        # Show completion
        self.progress_bar["value"] = 100
        self.progress_var.set(f"Done! Exported {total_slides} slide(s).")
        # Re-enable export button
        self.export_btn.config(state=tk.NORMAL)

        # Show appropriate dialog based on result
        if errors:
            messagebox.showerror(
                "Export errors",
                "Some files failed:\n\n" + "\n".join(errors),
            )
        elif total_slides > 0:
            messagebox.showinfo("Success", f"Exported {total_slides} slide(s) successfully.")

    def run(self):
        # Start the Tkinter event loop; blocks until window is closed
        self.root.mainloop()
