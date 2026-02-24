# excel_cleaner/cleaner/gui.py

# ---------------- ABOUT ----------------
# Author: Micah Braun
# AI Acknowledgement: This file was compiled with assistance from
#                     Copilot alongside Enterprise Data Protection.
# Date: 02/24/2026
'''
GUI module for the Classroom Utilization Cleaner application.

This module defines the graphical user interface (GUI) for running the
classroom utilization cleaning workflow defined in logic.py. It provides a
simple interface for selecting a raw EMS-exported CSV file, choosing where
to save the cleaned Excel output, and executing the cleaning process with a
single click.

Key features:
    - Input CSV file picker.
    - Output Excel file picker (defaulting to the user's Desktop).
    - Run button (enabled only when both input and output paths are chosen).
    - Cancel button to close the application.
    - Status bar that reflects the current state of the process.
    - Post-completion prompt allowing the user to process additional files.
'''
# ---------------- LIBRARIES ----------------
import os
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox

from .logic import run_cleaner

# ---------------- FUNCTIONS ----------------

from .logic import run_cleaner


def gui_main() -> None:
    '''
    Launch the Classroom Utilization Cleaner graphical user interface.

    This function initializes and displays a Tkinter-based GUI that provides
    users with a point and click interface for converting exported EMS classroom
    utilization CSV files into cleaned Excel reports. The GUI allows the user
    to browse for an input file, select an output destination (defaulting to
    the Desktop), initiate the cleaning, and optionally run multiple cleaning
    operations in succession.

    Interface features:
        - Input CSV selector
        - Output file selector (defaults to Desktop)
        - Run button (enabled only when both paths exist)
        - Cancel button to close the program
        - Status bar reflecting the current operation
        - Yes/No prompt after each successful run for multi-file workflows

    Args:
        None

    Returns:
        None
            The function starts the Tkinter event loop and does not return
            until the window is closed.

    Raises:
        None explicitly. Any exceptions that occur during event handling are
        caught and surfaced to the user in a Tkinter message box.
    '''

    root = tk.Tk()
    root.title('Classroom Utilization Cleaner')

    # -------------------------------
    # Styling and window defaults
    # -------------------------------
    bg_color = '#f4f4f4'
    accent_color = '#8B0000'

    root.configure(bg=bg_color)
    root.minsize(650, 220)

    def center_window(win: tk.Tk) -> None:
        '''
        Center the given Tkinter window on the user's screen.

        Args:
            win (tk.Tk):
                The Tk root or Toplevel window to reposition.

        Returns:
            None
        '''
        win.update_idletasks()
        width = win.winfo_width()
        height = win.winfo_height()
        screen_width = win.winfo_screenwidth()
        screen_height = win.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        win.geometry(f'{width}x{height}+{x}+{y}')

    # -------------------------------
    # Paths and Tkinter variables
    # -------------------------------
    user_home = os.path.expanduser('~')
    desktop_default = os.path.join(user_home, 'Desktop')
    if not os.path.isdir(desktop_default):
        desktop_default = user_home

    input_var = tk.StringVar()
    output_var = tk.StringVar()
    status_var = tk.StringVar(value='Select an input CSV and choose an output location.')

    # Main padded frame
    main_frame = tk.Frame(root, bg=bg_color, padx=15, pady=15)
    main_frame.grid(row=0, column=0, sticky='nsew')

    root.grid_rowconfigure(0, weight=1)
    root.grid_columnconfigure(0, weight=1)

    # Shared fonts
    default_font = ('Segoe UI', 10)
    label_font = ('Segoe UI', 10)
    status_font = ('Segoe UI', 9)

    # -------------------------------
    # Helper functions / Callbacks
    # -------------------------------

    def browse_input() -> None:
        '''
        Open a file-picker dialog for selecting the raw EMS CSV file.

        When a file is selected, this function:
            - Sets the input path
            - Suggests an output path on the Desktop using the same filename
              with a `_Clean.xlsx` suffix
            - Updates the status bar
        '''
        filename = filedialog.askopenfilename(
            title='Select EMS CSV Export',
            filetypes=[('CSV files', '*.csv'), ('All files', '*.*')]
        )
        if filename:
            input_var.set(filename)
            base = os.path.splitext(os.path.basename(filename))[0]
            output_var.set(os.path.join(desktop_default, base + '_Clean.xlsx'))
            status_var.set('Input selected. Confirm or adjust output location.')

    def browse_output() -> None:
        '''
        Open a save dialog to choose the output Excel file location.

        If the user already selected an input CSV, this function uses its
        basename when proposing the default output filename. Otherwise, a
        generic filename is suggested.
        '''
        input_path = input_var.get().strip()
        if input_path:
            base_name = os.path.splitext(os.path.basename(input_path))[0] + '_Clean.xlsx'
        else:
            base_name = 'Classroom_Utilization_Clean.xlsx'

        filename = filedialog.asksaveasfilename(
            title='Save Cleaned Excel As',
            initialdir=desktop_default,
            initialfile=base_name,
            defaultextension='.xlsx',
            filetypes=[('Excel files', '*.xlsx'), ('All files', '*.*')]
        )
        if filename:
            output_var.set(filename)
            status_var.set('Output location selected. Click Run to begin.')

    def reset_for_next_file() -> None:
        '''
        Reset the GUI fields to allow the user to clean another CSV file.

        Clears both input and output fields while keeping the window open
        for further processing.
        '''
        input_var.set('')
        output_var.set('')
        status_var.set('Select an input CSV and choose an output location.')

    def run_process() -> None:
        '''
        Execute the cleaning workflow for the selected input and output paths.

        This function:
            - Validates that both input and output paths are provided
            - Ensures the input path exists
            - Calls run_cleaner() from logic.py to perform the data transformation
            - Displays success or error dialogs
            - Prompts the user to clean another file or exit

        Returns:
            None
        '''
        input_path = input_var.get().strip()
        output_path = output_var.get().strip()

        if not input_path:
            messagebox.showerror('Error', 'Please select an input CSV file.')
            return
        if not os.path.isfile(input_path):
            messagebox.showerror('Error', f'Input file does not exist:\n{input_path}')
            return
        if not output_path:
            messagebox.showerror('Error', 'Please specify an output location.')
            return

        try:
            status_var.set('Running cleaning process...')
            root.update_idletasks()

            run_cleaner(input_path, output_path)

            status_var.set(f'Done! Saved to:\n{output_path}')
            messagebox.showinfo('Success', f'Cleaned report saved to:\n{output_path}')

            another = messagebox.askyesno(
                'Process another file?',
                'Do you want to clean another file?'
            )

            if another:
                reset_for_next_file()
            else:
                root.destroy()

        except Exception as exc:
            traceback.print_exc()
            messagebox.showerror('Error', f'An error occurred:\n{exc}')
            status_var.set('Error occurred — see message for details.')

    def cancel_and_close() -> None:
        '''
        Close the application gracefully by destroying the Tk root window.
        '''
        root.destroy()

    # -------------------------------
    # GUI Layout
    # -------------------------------

    title_label = tk.Label(
        main_frame,
        text='Classroom Utilization Cleaner',
        font=('Segoe UI', 12, 'bold'),
        bg=bg_color,
        fg=accent_color
    )
    title_label.grid(row=0, column=0, columnspan=3, sticky='w', pady=(0, 10))

    # Input selector
    tk.Label(main_frame, text='Input CSV:', font=label_font, bg=bg_color)\
        .grid(row=1, column=0, sticky='e', padx=5, pady=5)

    tk.Entry(main_frame, textvariable=input_var, width=60, font=default_font)\
        .grid(row=1, column=1, padx=5, pady=5, sticky='we')

    tk.Button(main_frame, text='Browse...', command=browse_input, font=default_font)\
        .grid(row=1, column=2, padx=5, pady=5)

    # Output selector
    tk.Label(main_frame, text='Output Excel:', font=label_font, bg=bg_color)\
        .grid(row=2, column=0, sticky='e', padx=5, pady=5)

    tk.Entry(main_frame, textvariable=output_var, width=60, font=default_font)\
        .grid(row=2, column=1, padx=5, pady=5, sticky='we')

    tk.Button(main_frame, text='Browse...', command=browse_output, font=default_font)\
        .grid(row=2, column=2, padx=5, pady=5)

    # Buttons
    button_frame = tk.Frame(main_frame, bg=bg_color)
    button_frame.grid(row=3, column=0, columnspan=3, pady=(10, 5))

    run_button = tk.Button(
            button_frame,
            text='Run',
            width=12,
            command=run_process,
            state=tk.DISABLED,
            bg=accent_color,
            fg='white',
            activebackground='#a80000',
            activeforeground='white',
            font=default_font
        )
    run_button.pack(side='left', padx=10)

    tk.Button(
            button_frame,
            text='Cancel',
            width=12,
            command=cancel_and_close,
            font=default_font
        ).pack(side='left', padx=10)

        # Status bar
    status_label = tk.Label(
            main_frame,
            textvariable=status_var,
            font=status_font,
            bg='#e6e6e6',
            fg='#333333',
            anchor='w',
            padx=8,
            pady=4,
            relief='sunken'
        )
    status_label.grid(row=4, column=0, columnspan=3, sticky='we', pady=(10, 0))

    main_frame.grid_columnconfigure(1, weight=1)

    # Enable/disable Run button
    def update_run_button_state(*_: object) -> None:
        '''
        Enable or disable the Run button depending on whether both
        input and output paths have been provided by the user.
        '''
        if input_var.get().strip() and output_var.get().strip():
            run_button.config(state=tk.NORMAL)
        else:
            run_button.config(state=tk.DISABLED)

    input_var.trace_add('write', update_run_button_state)
    output_var.trace_add('write', update_run_button_state)
    update_run_button_state()

    center_window(root)
    root.resizable(False, False)
    root.mainloop()