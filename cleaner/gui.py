# excel_cleaner/cleaner/gui.py

# ---------------- ABOUT ----------------
# Author: Micah Braun
# AI Acknowledgement: This file was compiled with assistance from
#                     Copilot alongside Enterprise Data Protection.
# Date: 02/24/2026
# Last Updated: 03/03/2026
"""
GUI module for the Classroom Utilization Cleaner application with
support for single-file and batch modes.

- Single mode:
    * Select one EMS export file (.xls, .xlsx, .csv)
    * Select a single output Excel file
    * Uses run_cleaner()

- Batch mode:
    * Select multiple input files and/or a folder
    * Select an output folder
    * Uses run_batch_cleaner()
"""

import os
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox

from .logic import run_cleaner, run_batch_cleaner, ALLOWED_EXTENSIONS

# excel_cleaner/cleaner/gui.py

import os
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox

from .logic import run_cleaner, run_batch_cleaner, ALLOWED_EXTENSIONS


def create_app():
    """
    Build the Tkinter UI but DO NOT start the mainloop.
    Returns:
        root: the Tk root window
        ctx:  a dict containing key state & handlers for testing:
              {
                  'mode_var': ...,
                  'input_var': ...,
                  'output_var': ...,
                  'status_var': ...,
                  'run_process': <callable>,
              }
    """
    root = tk.Tk()
    root.title("Classroom Utilization Cleaner")

    bg_color = "#f4f4f4"
    accent_color = "#8B0000"
    root.configure(bg=bg_color)
    root.minsize(700, 280)

    def center_window(win: tk.Tk) -> None:
        win.update_idletasks()
        width = win.winfo_width()
        height = win.winfo_height()
        screen_width = win.winfo_screenwidth()
        screen_height = win.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        win.geometry(f"{width}x{height}+{x}+{y}")

    # Tk variables
    mode_var = tk.StringVar(value="single")  # "single" or "batch"
    input_var = tk.StringVar()
    output_var = tk.StringVar()
    status_var = tk.StringVar(value="Select mode, then choose input and output.")

    user_home = os.path.expanduser("~")
    desktop_default = os.path.join(user_home, "Desktop")
    if not os.path.isdir(desktop_default):
        desktop_default = user_home

    main_frame = tk.Frame(root, bg=bg_color, padx=15, pady=15)
    main_frame.grid(row=0, column=0, sticky="nsew")
    root.grid_rowconfigure(0, weight=1)
    root.grid_columnconfigure(0, weight=1)

    default_font = ("Segoe UI", 10)
    label_font = ("Segoe UI", 10)
    status_font = ("Segoe UI", 9)

    # Title
    title_label = tk.Label(
        main_frame,
        text="Classroom Utilization Cleaner",
        font=("Segoe UI", 12, "bold"),
        bg=bg_color,
        fg=accent_color,
    )
    title_label.grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 5))

    # Warning label
    warning_label = tk.Label(
        main_frame,
        text="Only .xls, .xlsx, or .csv files can be used.",
        font=("Segoe UI", 9, "italic"),
        fg="red",
        bg=bg_color,
    )
    warning_label.grid(row=1, column=0, columnspan=4, sticky="w", pady=(0, 10))

    # Mode toggle
    mode_frame = tk.Frame(main_frame, bg=bg_color)
    mode_frame.grid(row=2, column=0, columnspan=4, sticky="w", pady=(0, 10))

    tk.Label(mode_frame, text="Mode:", font=label_font, bg=bg_color).pack(side="left", padx=(0, 10))

    def set_single_mode():
        status_var.set("Single mode: choose one input file and one output file.")

    def set_batch_mode():
        status_var.set("Batch mode: choose files or a folder and an output folder.")

    tk.Radiobutton(
        mode_frame,
        text="Single file",
        variable=mode_var,
        value="single",
        font=default_font,
        bg=bg_color,
        command=set_single_mode,
    ).pack(side="left", padx=5)

    tk.Radiobutton(
        mode_frame,
        text="Batch (multiple files / folder)",
        variable=mode_var,
        value="batch",
        font=default_font,
        bg=bg_color,
        command=set_batch_mode,
    ).pack(side="left", padx=5)

    # Input browse
    def browse_input() -> None:
        mode = mode_var.get()

        if mode == "single":
            filename = filedialog.askopenfilename(
                title="Select EMS Exported File",
                filetypes=[
                    ("Allowed Files", "*.xls *.xlsx *.csv"),
                    ("All files", "*.*"),
                ],
            )
            if filename:
                input_var.set(filename)
                base = os.path.splitext(os.path.basename(filename))[0]
                output_var.set(os.path.join(desktop_default, base + "_Clean.xlsx"))
                status_var.set("Single file selected. Confirm or choose output Excel file.")
        else:
            choice = messagebox.askyesno(
                "Batch input",
                "Click 'Yes' to select multiple files.\n"
                "Click 'No' to select an entire folder."
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
                    input_var.set(", ".join(filenames))
                    output_var.set(desktop_default)
                    status_var.set("Multiple files selected. Choose output folder.")
            else:
                folder = filedialog.askdirectory(
                    title="Select Folder Containing EMS Files"
                )
                if folder:
                    input_var.set(folder)
                    output_var.set(desktop_default)
                    status_var.set("Folder selected. Choose output folder.")

    # Output browse
    def browse_output() -> None:
        mode = mode_var.get()
        raw_input_val = input_var.get().strip()

        if not raw_input_val:
            messagebox.showerror("Error", "Please select input first.")
            return

        if mode == "single":
            base_name = "Classroom_Utilization_Clean.xlsx"
            if raw_input_val:
                base_name = os.path.splitext(os.path.basename(raw_input_val))[0] + "_Clean.xlsx"

            filename = filedialog.asksaveasfilename(
                title="Save Cleaned Excel As",
                initialdir=desktop_default,
                initialfile=base_name,
                defaultextension=".xlsx",
                filetypes=[("Excel Files", "*.xlsx")],
            )
            if filename:
                output_var.set(filename)
                status_var.set("Output file selected. Click Run to begin.")
        else:
            folder = filedialog.askdirectory(
                title="Select Output Folder",
                initialdir=desktop_default,
            )
            if folder:
                output_var.set(folder)
                status_var.set("Output folder selected. Click Run to begin.")

    def reset_for_next() -> None:
        input_var.set("")
        output_var.set("")
        if mode_var.get() == "single":
            set_single_mode()
        else:
            set_batch_mode()

    def run_process() -> None:
        mode = mode_var.get()
        raw_input_val = input_var.get().strip()
        output_path = output_var.get().strip()

        if not raw_input_val:
            messagebox.showerror("Error", "Please select input.")
            return
        if not output_path:
            messagebox.showerror("Error", "Please select an output location.")
            return

        try:
            status_var.set("Running …")
            root.update_idletasks()

            if mode == "single":
                if not os.path.isfile(raw_input_val):
                    messagebox.showerror("Error", f"Input file does not exist:\n{raw_input_val}")
                    status_var.set("Error — input file not found.")
                    return

                run_cleaner(raw_input_val, output_path)
                messagebox.showinfo("Success", f"Cleaned report saved to:\n{output_path}")
            else:
                if os.path.isdir(raw_input_val):
                    inputs = [raw_input_val]
                else:
                    inputs = [x.strip() for x in raw_input_val.split(",") if x.strip()]

                run_batch_cleaner(inputs, output_path)
                messagebox.showinfo("Success", "Batch processing complete.")

            another = messagebox.askyesno(
                "Process another?",
                "Do you want to run another cleaning operation?"
            )
            if another:
                reset_for_next()
            else:
                root.destroy()

        except Exception as exc:
            traceback.print_exc()
            messagebox.showerror("Error", f"An error occurred:\n{exc}")
            status_var.set("Error — see message for details.")

    def cancel_and_close() -> None:
        root.destroy()

    # Input row
    tk.Label(
        main_frame,
        text="Input:",
        font=label_font,
        bg=bg_color,
    ).grid(row=3, column=0, sticky="e", padx=5, pady=5)

    tk.Entry(
        main_frame,
        textvariable=input_var,
        width=60,
        font=default_font,
    ).grid(row=3, column=1, padx=5, pady=5, sticky="we")

    tk.Button(
        main_frame,
        text="Browse…",
        command=browse_input,
        font=default_font,
    ).grid(row=3, column=2, padx=5, pady=5, sticky="w")

    # Output row
    tk.Label(
        main_frame,
        text="Output:",
        font=label_font,
        bg=bg_color,
    ).grid(row=4, column=0, sticky="e", padx=5, pady=5)

    tk.Entry(
        main_frame,
        textvariable=output_var,
        width=60,
        font=default_font,
    ).grid(row=4, column=1, padx=5, pady=5, sticky="we")

    tk.Button(
        main_frame,
        text="Browse…",
        command=browse_output,
        font=default_font,
    ).grid(row=4, column=2, padx=5, pady=5, sticky="w")

    main_frame.grid_columnconfigure(1, weight=1)

    # Buttons
    button_frame = tk.Frame(main_frame, bg=bg_color)
    button_frame.grid(row=5, column=0, columnspan=3, pady=(10, 5))

    run_button = tk.Button(
        button_frame,
        text="Run",
        width=12,
        command=run_process,
        bg=accent_color,
        fg="white",
        activebackground="#a80000",
        activeforeground="white",
        font=default_font,
    )
    run_button.pack(side="left", padx=10)

    tk.Button(
        button_frame,
        text="Cancel",
        width=12,
        command=cancel_and_close,
        font=default_font,
    ).pack(side="left", padx=10)

    # Status bar
    status_label = tk.Label(
        main_frame,
        textvariable=status_var,
        font=status_font,
        bg="#e6e6e6",
        fg="#333333",
        anchor="w",
        padx=8,
        pady=4,
        relief="sunken",
    )
    status_label.grid(row=6, column=0, columnspan=3, sticky="we", pady=(10, 0))

    center_window(root)
    root.resizable(False, False)

    ctx = {
        "mode_var": mode_var,
        "input_var": input_var,
        "output_var": output_var,
        "status_var": status_var,
        "run_process": run_process,
    }
    return root, ctx


def gui_main() -> None:
    """Production entrypoint: build the app and start the Tk mainloop."""
    root, _ = create_app()
    root.mainloop()

def gui_main() -> None:
    root = tk.Tk()
    root.title("Classroom Utilization Cleaner")

    # -------------------------------
    # Styling and window defaults
    # -------------------------------
    bg_color = "#f4f4f4"
    accent_color = "#8B0000"
    root.configure(bg=bg_color)
    root.minsize(700, 280)

    def center_window(win: tk.Tk) -> None:
        win.update_idletasks()
        width = win.winfo_width()
        height = win.winfo_height()
        screen_width = win.winfo_screenwidth()
        screen_height = win.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        win.geometry(f"{width}x{height}+{x}+{y}")

    # -------------------------------
    # Tk variables
    # -------------------------------
    mode_var = tk.StringVar(value="single")  # "single" or "batch"
    input_var = tk.StringVar()
    output_var = tk.StringVar()
    status_var = tk.StringVar(value="Select mode, then choose input and output.")

    user_home = os.path.expanduser("~")
    desktop_default = os.path.join(user_home, "Desktop")
    if not os.path.isdir(desktop_default):
        desktop_default = user_home

    # -------------------------------
    # Layout - main frame
    # -------------------------------
    main_frame = tk.Frame(root, bg=bg_color, padx=15, pady=15)
    main_frame.grid(row=0, column=0, sticky="nsew")
    root.grid_rowconfigure(0, weight=1)
    root.grid_columnconfigure(0, weight=1)

    default_font = ("Segoe UI", 10)
    label_font = ("Segoe UI", 10)
    status_font = ("Segoe UI", 9)

    # -------------------------------
    # Title
    # -------------------------------
    title_label = tk.Label(
        main_frame,
        text="Classroom Utilization Cleaner",
        font=("Segoe UI", 12, "bold"),
        bg=bg_color,
        fg=accent_color,
    )
    title_label.grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 5))

    # -------------------------------
    # Warning label
    # -------------------------------
    warning_label = tk.Label(
        main_frame,
        text="Only .xls, .xlsx, or .csv files can be used.",
        font=("Segoe UI", 9, "italic"),
        fg="red",
        bg=bg_color,
    )
    warning_label.grid(row=1, column=0, columnspan=4, sticky="w", pady=(0, 10))

    # -------------------------------
    # Mode toggle (Single vs Batch)
    # -------------------------------
    mode_frame = tk.Frame(main_frame, bg=bg_color)
    mode_frame.grid(row=2, column=0, columnspan=4, sticky="w", pady=(0, 10))

    tk.Label(
        mode_frame,
        text="Mode:",
        font=label_font,
        bg=bg_color,
    ).pack(side="left", padx=(0, 10))

    tk.Radiobutton(
        mode_frame,
        text="Single file",
        variable=mode_var,
        value="single",
        font=default_font,
        bg=bg_color,
        command=lambda: status_var.set("Single mode: choose one input file and one output file."),
    ).pack(side="left", padx=5)

    tk.Radiobutton(
        mode_frame,
        text="Batch (multiple files / folder)",
        variable=mode_var,
        value="batch",
        font=default_font,
        bg=bg_color,
        command=lambda: status_var.set("Batch mode: choose files or a folder and an output folder."),
    ).pack(side="left", padx=5)

    # -------------------------------
    # Browse Input
    # -------------------------------
    def browse_input() -> None:
        mode = mode_var.get()

        if mode == "single":
            # Single-file mode: one file
            filename = filedialog.askopenfilename(
                title="Select EMS Exported File",
                filetypes=[
                    ("Allowed Files", "*.xls *.xlsx *.csv"),
                    ("All files", "*.*"),
                ],
            )
            if filename:
                input_var.set(filename)
                base = os.path.splitext(os.path.basename(filename))[0]
                # Suggest output file on Desktop
                output_var.set(os.path.join(desktop_default, base + "_Clean.xlsx"))
                status_var.set("Single file selected. Confirm or choose output Excel file.")
        else:
            # Batch mode: multi-file or folder
            # Let user choose either multiple files or a folder via a small dialog
            choice = messagebox.askyesno(
                "Batch input",
                "Click 'Yes' to select multiple files.\n"
                "Click 'No' to select an entire folder."
            )
            if choice:
                # Multiple files
                filenames = filedialog.askopenfilenames(
                    title="Select EMS Files (.xls, .xlsx, .csv)",
                    filetypes=[
                        ("Allowed Files", "*.xls *.xlsx *.csv"),
                        ("All files", "*.*"),
                    ],
                )
                if filenames:
                    input_var.set(", ".join(filenames))
                    # Suggest output folder as Desktop
                    output_var.set(desktop_default)
                    status_var.set("Multiple files selected. Choose output folder.")
            else:
                # Folder
                folder = filedialog.askdirectory(
                    title="Select Folder Containing EMS Files"
                )
                if folder:
                    input_var.set(folder)
                    output_var.set(desktop_default)
                    status_var.set("Folder selected. Choose output folder.")

    # -------------------------------
    # Browse Output
    # -------------------------------
    def browse_output() -> None:
        mode = mode_var.get()
        raw_input_val = input_var.get().strip()

        if not raw_input_val:
            messagebox.showerror("Error", "Please select input first.")
            return

        if mode == "single":
            # Output is a single Excel file
            base_name = "Classroom_Utilization_Clean.xlsx"
            if raw_input_val:
                base_name = os.path.splitext(os.path.basename(raw_input_val))[0] + "_Clean.xlsx"

            filename = filedialog.asksaveasfilename(
                title="Save Cleaned Excel As",
                initialdir=desktop_default,
                initialfile=base_name,
                defaultextension=".xlsx",
                filetypes=[("Excel Files", "*.xlsx")],
            )
            if filename:
                output_var.set(filename)
                status_var.set("Output file selected. Click Run to begin.")
        else:
            # Batch mode: output is a folder
            folder = filedialog.askdirectory(
                title="Select Output Folder",
                initialdir=desktop_default,
            )
            if folder:
                output_var.set(folder)
                status_var.set("Output folder selected. Click Run to begin.")

    # -------------------------------
    # Reset
    # -------------------------------
    def reset_for_next() -> None:
        input_var.set("")
        output_var.set("")
        if mode_var.get() == "single":
            status_var.set("Single mode: choose one input file and one output file.")
        else:
            status_var.set("Batch mode: choose files or a folder and an output folder.")

    # -------------------------------
    # Run process
    # -------------------------------
    def run_process() -> None:
        mode = mode_var.get()
        raw_input_val = input_var.get().strip()
        output_path = output_var.get().strip()

        if not raw_input_val:
            messagebox.showerror("Error", "Please select input.")
            return
        if not output_path:
            messagebox.showerror("Error", "Please select an output location.")
            return

        try:
            status_var.set("Running …")
            root.update_idletasks()

            if mode == "single":
                # Validate single input file exists
                if not os.path.isfile(raw_input_val):
                    messagebox.showerror("Error", f"Input file does not exist:\n{raw_input_val}")
                    status_var.set("Error — input file not found.")
                    return

                run_cleaner(raw_input_val, output_path)
                messagebox.showinfo("Success", f"Cleaned report saved to:\n{output_path}")
            else:
                # Batch mode: raw_input_val can be a folder or comma-separated list of files
                if os.path.isdir(raw_input_val):
                    inputs = [raw_input_val]
                else:
                    inputs = [x.strip() for x in raw_input_val.split(",") if x.strip()]

                run_batch_cleaner(inputs, output_path)
                messagebox.showinfo("Success", "Batch processing complete.")

            # Ask to process more
            another = messagebox.askyesno(
                "Process another?",
                "Do you want to run another cleaning operation?"
            )
            if another:
                reset_for_next()
            else:
                root.destroy()

        except Exception as exc:
            traceback.print_exc()
            messagebox.showerror("Error", f"An error occurred:\n{exc}")
            status_var.set("Error — see message for details.")

    # -------------------------------
    # Cancel
    # -------------------------------
    def cancel_and_close() -> None:
        root.destroy()

    # -------------------------------
    # Widgets - Input row
    # -------------------------------
    tk.Label(
        main_frame,
        text="Input:",
        font=label_font,
        bg=bg_color,
    ).grid(row=3, column=0, sticky="e", padx=5, pady=5)

    tk.Entry(
        main_frame,
        textvariable=input_var,
        width=60,
        font=default_font,
    ).grid(row=3, column=1, padx=5, pady=5, sticky="we")

    tk.Button(
        main_frame,
        text="Browse…",
        command=browse_input,
        font=default_font,
    ).grid(row=3, column=2, padx=5, pady=5, sticky="w")

    # -------------------------------
    # Widgets - Output row
    # -------------------------------
    tk.Label(
        main_frame,
        text="Output:",
        font=label_font,
        bg=bg_color,
    ).grid(row=4, column=0, sticky="e", padx=5, pady=5)

    tk.Entry(
        main_frame,
        textvariable=output_var,
        width=60,
        font=default_font,
    ).grid(row=4, column=1, padx=5, pady=5, sticky="we")

    tk.Button(
        main_frame,
        text="Browse…",
        command=browse_output,
        font=default_font,
    ).grid(row=4, column=2, padx=5, pady=5, sticky="w")

    main_frame.grid_columnconfigure(1, weight=1)

    # -------------------------------
    # Buttons
    # -------------------------------
    button_frame = tk.Frame(main_frame, bg=bg_color)
    button_frame.grid(row=5, column=0, columnspan=3, pady=(10, 5))

    run_button = tk.Button(
        button_frame,
        text="Run",
        width=12,
        command=run_process,
        bg=accent_color,
        fg="white",
        activebackground="#a80000",
        activeforeground="white",
        font=default_font,
    )
    run_button.pack(side="left", padx=10)

    tk.Button(
        button_frame,
        text="Cancel",
        width=12,
        command=cancel_and_close,
        font=default_font,
    ).pack(side="left", padx=10)

    # -------------------------------
    # Status bar
    # -------------------------------
    status_label = tk.Label(
        main_frame,
        textvariable=status_var,
        font=status_font,
        bg="#e6e6e6",
        fg="#333333",
        anchor="w",
        padx=8,
        pady=4,
        relief="sunken",
    )
    status_label.grid(row=6, column=0, columnspan=3, sticky="we", pady=(10, 0))

    center_window(root)
    root.resizable(False, False)
    root.mainloop()