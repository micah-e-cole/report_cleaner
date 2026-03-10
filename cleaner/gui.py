# excel_cleaner/cleaner/gui.py
# ------------- ABOUT -------------
# Author: Micah Braun
# Two-page wizard GUI:
#   Page 1: choose report type
#   Page 2: choose single/batch, input/output, run

import os
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

from .dispatcher import run_cleaner, run_batch_cleaner, REPORT_TYPES
from .common import ALLOWED_EXTENSIONS, collect_input_files

BG_COLOR = "#f4f4f4"
ACCENT_COLOR = "#8B0000"


def center_window(win: tk.Tk) -> None:
    """Center the window on the screen."""
    win.update_idletasks()
    width = win.winfo_width()
    height = win.winfo_height()
    screen_width = win.winfo_screenwidth()
    screen_height = win.winfo_screenheight()
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2
    win.geometry(f"{width}x{height}+{x}+{y}")


# ----------------------------------------------------------------------
# Page 1: Report type selection
# ----------------------------------------------------------------------
class ReportTypeSelectorFrame(tk.Frame):
    """
    First page of the wizard: ask which report type the user is cleaning.
    After selection, transitions to CleaningOptionsFrame.
    """

    def __init__(self, parent: tk.Widget, root: tk.Tk):
        super().__init__(parent, bg=BG_COLOR)
        self.root = root
        self.parent = parent

        title_font = ("Segoe UI", 14, "bold")
        label_font = ("Segoe UI", 10)

        # Title
        tk.Label(
            self,
            text="EMS File Cleaner",
            font=title_font,
            bg=BG_COLOR,
            fg=ACCENT_COLOR,
        ).pack(pady=(20, 5))

        # Subtitle
        tk.Label(
            self,
            text="Step 1: Choose the type of report you want to clean.",
            font=label_font,
            bg=BG_COLOR,
        ).pack(pady=(0, 15))

        # Warning / info
        tk.Label(
            self,
            text=f"Only {', '.join(sorted(ALLOWED_EXTENSIONS))} files can be used.",
            font=("Segoe UI", 9, "italic"),
            fg="red",
            bg=BG_COLOR,
        ).pack(pady=(0, 20))

        # Report type dropdown
        self.report_var = tk.StringVar(
            value=list(REPORT_TYPES.keys())[0] if REPORT_TYPES else ""
        )

        tk.Label(
            self,
            text="Report Type:",
            font=label_font,
            bg=BG_COLOR,
        ).pack(pady=(0, 5))

        if REPORT_TYPES:
            report_dropdown = ttk.Combobox(
                self,
                textvariable=self.report_var,
                values=list(REPORT_TYPES.keys()),
                state="readonly",
                font=("Segoe UI", 10),
            )
            report_dropdown.pack(pady=(0, 20), ipadx=10)

        else:
            tk.Label(
                self,
                text="No report types are registered.",
                font=("Segoe UI", 10, "italic"),
                fg="red",
                bg=BG_COLOR,
            ).pack(pady=(0, 20))

        # Buttons
        button_frame = tk.Frame(self, bg=BG_COLOR)
        button_frame.pack(pady=(10, 10))

        tk.Button(
            button_frame,
            text="Next →",
            width=14,
            command=self.go_next,
            bg=ACCENT_COLOR,
            fg="white",
            activebackground="#a80000",
            activeforeground="white",
            font=("Segoe UI", 10),
        ).pack(side="left", padx=10)

        tk.Button(
            button_frame,
            text="Cancel",
            width=14,
            command=root.destroy,
            font=("Segoe UI", 10),
        ).pack(side="left", padx=10)

    def go_next(self) -> None:
        report_type = self.report_var.get()
        if not report_type:
            messagebox.showerror("Error", "No report types available.")
            return

        # Destroy this page, show the next
        self.destroy()
        CleaningOptionsFrame(self.parent, self.root, report_type).pack(
            fill="both", expand=True
        )


# ----------------------------------------------------------------------
# Page 2: Cleaning options (single vs batch for the chosen report type)
# ----------------------------------------------------------------------
class CleaningOptionsFrame(tk.Frame):
    """
    Second page of the wizard: for a given report type, choose:
    - Single vs Batch
    - Input path(s)
    - Output path or folder
    - Run the process
    """

    def __init__(self, parent: tk.Widget, root: tk.Tk, report_type: str):
        super().__init__(parent, bg=BG_COLOR)
        self.root = root
        self.parent = parent
        self.report_type = report_type

        # Fonts
        title_font = ("Segoe UI", 14, "bold")
        default_font = ("Segoe UI", 10)
        label_font = ("Segoe UI", 10)
        status_font = ("Segoe UI", 9)

        # Grid layout config
        self.grid_rowconfigure(0, weight=0)
        self.grid_rowconfigure(1, weight=0)
        self.grid_rowconfigure(2, weight=0)
        self.grid_rowconfigure(3, weight=0)
        self.grid_rowconfigure(4, weight=0)
        self.grid_rowconfigure(5, weight=1)
        self.grid_columnconfigure(1, weight=1)

        # Tk variables
        self.mode_var = tk.StringVar(value="single")  # "single" or "batch"
        self.input_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.status_var = tk.StringVar(
            value=f"{report_type}: choose mode, then input and output."
        )

        user_home = os.path.expanduser("~")
        self.desktop_default = os.path.join(user_home, "Desktop")
        if not os.path.isdir(self.desktop_default):
            self.desktop_default = user_home

        # Title
        tk.Label(
            self,
            text=f"{report_type} Cleaner",
            font=title_font,
            bg=BG_COLOR,
            fg=ACCENT_COLOR,
        ).grid(row=0, column=0, columnspan=4, sticky="w", pady=(10, 5), padx=15)

        # Warning label
        tk.Label(
            self,
            text=f"Only {', '.join(sorted(ALLOWED_EXTENSIONS))} files can be used.",
            font=("Segoe UI", 9, "italic"),
            fg="red",
            bg=BG_COLOR,
        ).grid(row=1, column=0, columnspan=4, sticky="w", padx=15, pady=(0, 10))

        # Mode toggle
        mode_frame = tk.Frame(self, bg=BG_COLOR)
        mode_frame.grid(row=2, column=0, columnspan=4, sticky="w", padx=15, pady=(0, 10))

        tk.Label(
            mode_frame, text="Mode:", font=label_font, bg=BG_COLOR
        ).pack(side="left", padx=(0, 10))

        tk.Radiobutton(
            mode_frame,
            text="Single file",
            variable=self.mode_var,
            value="single",
            font=default_font,
            bg=BG_COLOR,
            command=lambda: self.status_var.set(
                f"{self.report_type}: single mode — choose one input file and one output file."
            ),
        ).pack(side="left", padx=5)

        tk.Radiobutton(
            mode_frame,
            text="Batch (multiple files / folder)",
            variable=self.mode_var,
            value="batch",
            font=default_font,
            bg=BG_COLOR,
            command=lambda: self.status_var.set(
                f"{self.report_type}: batch mode — choose files or a folder and an output folder."
            ),
        ).pack(side="left", padx=5)

        # Input row
        tk.Label(
            self,
            text="Input:",
            font=label_font,
            bg=BG_COLOR,
        ).grid(row=3, column=0, sticky="e", padx=10, pady=5)

        tk.Entry(
            self,
            textvariable=self.input_var,
            width=60,
            font=default_font,
        ).grid(row=3, column=1, padx=5, pady=5, sticky="we")

        tk.Button(
            self,
            text="Browse…",
            command=self.browse_input,
            font=default_font,
        ).grid(row=3, column=2, padx=5, pady=5, sticky="w")

        # Output row
        tk.Label(
            self,
            text="Output:",
            font=label_font,
            bg=BG_COLOR,
        ).grid(row=4, column=0, sticky="e", padx=10, pady=5)

        tk.Entry(
            self,
            textvariable=self.output_var,
            width=60,
            font=default_font,
        ).grid(row=4, column=1, padx=5, pady=5, sticky="we")

        tk.Button(
            self,
            text="Browse…",
            command=self.browse_output,
            font=default_font,
        ).grid(row=4, column=2, padx=5, pady=5, sticky="w")

        # Buttons
        button_frame = tk.Frame(self, bg=BG_COLOR)
        button_frame.grid(row=5, column=0, columnspan=3, pady=(10, 5))

        tk.Button(
            button_frame,
            text="← Back",
            width=12,
            command=self.go_back,
            font=default_font,
        ).pack(side="left", padx=10)

        tk.Button(
            button_frame,
            text="Run",
            width=12,
            command=self.run_process,
            bg=ACCENT_COLOR,
            fg="white",
            activebackground="#a80000",
            activeforeground="white",
            font=default_font,
        ).pack(side="left", padx=10)

        tk.Button(
            button_frame,
            text="Cancel",
            width=12,
            command=self.root.destroy,
            font=default_font,
        ).pack(side="left", padx=10)

        # Status bar
        status_label = tk.Label(
            self,
            textvariable=self.status_var,
            font=status_font,
            bg="#e6e6e6",
            fg="#333333",
            anchor="w",
            padx=8,
            pady=4,
            relief="sunken",
        )
        status_label.grid(
            row=6, column=0, columnspan=3, sticky="we", padx=10, pady=(10, 0)
        )

    # ---------------------- Navigation ---------------------- #
    def go_back(self) -> None:
        """Return to the report-type selection page."""
        self.destroy()
        ReportTypeSelectorFrame(self.parent, self.root).pack(fill="both", expand=True)

    # ---------------------- Browse Handlers ---------------------- #
    def browse_input(self) -> None:
        """Pick input file(s) or folder depending on mode."""
        mode = self.mode_var.get()

        if mode == "single":
            filename = filedialog.askopenfilename(
                title="Select EMS Exported File",
                filetypes=[
                    ("Allowed Files", "*.xls *.xlsx *.csv"),
                    ("All files", "*.*"),
                ],
            )
            if filename:
                self.input_var.set(filename)
                base = os.path.splitext(os.path.basename(filename))[0]
                self.output_var.set(os.path.join(self.desktop_default, base + "_Clean.xlsx"))
                self.status_var.set(
                    f"{self.report_type}: single file selected. Confirm or choose output Excel file."
                )
        else:
            choice = messagebox.askyesno(
                "Batch input",
                "Click 'Yes' to select multiple files.\n"
                "Click 'No' to select an entire folder.",
            )
            if choice:
                filenames = filedialog.askopenfilenames(
                    title="Select EMS Files (.xls, .xlsx, .csv)",
                    filetypes=[
                        ("Allowed Files", "*.xls *.xlsx *.csv"),
                        ("All files", "*.*"),
                    ],
                )
                if filenames:
                    self.input_var.set(", ".join(filenames))
                    self.output_var.set(self.desktop_default)
                    self.status_var.set(
                        f"{self.report_type}: multiple files selected. Choose output folder."
                    )
            else:
                folder = filedialog.askdirectory(
                    title="Select Folder Containing EMS Files"
                )
                if folder:
                    self.input_var.set(folder)
                    self.output_var.set(self.desktop_default)
                    self.status_var.set(
                        f"{self.report_type}: folder selected. Choose output folder."
                    )

    def browse_output(self) -> None:
        """Pick output file (single) or folder (batch)."""
        mode = self.mode_var.get()
        raw_input_val = self.input_var.get().strip()

        if not raw_input_val:
            messagebox.showerror("Error", "Please select input first.")
            return

        if mode == "single":
            base_name = f"{self.report_type.replace(' ', '_')}_Clean.xlsx"
            if raw_input_val:
                base_name = (
                    os.path.splitext(os.path.basename(raw_input_val))[0] + "_Clean.xlsx"
                )

            filename = filedialog.asksaveasfilename(
                title="Save Cleaned Excel As",
                initialdir=self.desktop_default,
                initialfile=base_name,
                defaultextension=".xlsx",
                filetypes=[("Excel Files", "*.xlsx")],
            )
            if filename:
                self.output_var.set(filename)
                self.status_var.set("Output file selected. Click Run to begin.")
        else:
            folder = filedialog.askdirectory(
                title="Select Output Folder",
                initialdir=self.desktop_default,
            )
            if folder:
                self.output_var.set(folder)
                self.status_var.set("Output folder selected. Click Run to begin.")

    # ---------------------- Run Handler ---------------------- #
    def run_process(self) -> None:
        """Execute the cleaning process for the chosen report type."""
        mode = self.mode_var.get()
        raw_input_val = self.input_var.get().strip()
        output_path = self.output_var.get().strip()
        report_type = self.report_type

        if not raw_input_val:
            messagebox.showerror("Error", "Please select input.")
            return
        if not output_path:
            messagebox.showerror("Error", "Please select an output location.")
            return

        try:
            self.status_var.set("Running …")
            self.root.update_idletasks()

            if mode == "single":
                if not os.path.isfile(raw_input_val):
                    messagebox.showerror(
                        "Error",
                        f"Input file does not exist:\n{raw_input_val}",
                    )
                    self.status_var.set("Error — input file not found.")
                    return

                run_cleaner(raw_input_val, output_path, report_type)
                messagebox.showinfo(
                    "Success", f"Cleaned report saved to:\n{output_path}"
                )

            else:
                # Batch mode
                if os.path.isdir(raw_input_val):
                    inputs = [raw_input_val]
                else:
                    inputs = [
                        x.strip()
                        for x in raw_input_val.split(",")
                        if x.strip()
                    ]

                run_batch_cleaner(
                    inputs, output_path, report_type, collect_input_files
                )
                messagebox.showinfo("Success", "Batch processing complete.")

            # Ask to process another
            another = messagebox.askyesno(
                "Process another?",
                "Do you want to run another cleaning operation?",
            )
            if another:
                # Return to first page
                self.go_back()
            else:
                self.root.destroy()

        except Exception as exc:
            traceback.print_exc()
            messagebox.showerror("Error", f"An error occurred:\n{exc}")
            self.status_var.set("Error — see message for details.")


# ----------------------------------------------------------------------
# Top-level entrypoint used by run_gui.py
# ----------------------------------------------------------------------
def gui_main() -> None:
    """
    Production entrypoint: build the two-page wizard and start Tk mainloop.
    """
    root = tk.Tk()
    root.title("Classroom Utilization Cleaner")
    root.configure(bg=BG_COLOR)
    root.minsize(700, 320)

    container = tk.Frame(root, bg=BG_COLOR)
    container.pack(fill="both", expand=True)

    first_page = ReportTypeSelectorFrame(container, root)
    first_page.pack(fill="both", expand=True)

    center_window(root)
    root.resizable(False, False)
    root.mainloop()


def create_app_for_tests():
    """ 
    Build the GUI without calling mainloop().
    Used only by automated tests"""
    root = tk.Tk()
    root.title("Classroom Utilization Cleaner (test)")
    root.configure(bg=BG_COLOR)
    root.minsize(700, 320)

    container = tk.Frame(root, bg=BG_COLOR)
    container.pack(fill="both", expand=True)

    first_page = ReportTypeSelectorFrame(container, root)
    first_page.pack(fill="both", expand=True)

    # Collect references needed by GUI tests
    ctx = {
        "report_var": first_page.report_var,
        # Not on this page yet, but test_gui.py switches pages — so:
        # These will be populated after navigating to CleaningOptionsFrame
    } 
    return root, ctx