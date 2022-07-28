import os
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Callable


# class for the GUI app
class GUIApp:
    def __init__(self, master: tk.Tk, run_callback: Callable[..., None]):

        self.run_callback = run_callback

        # input file path
        self.input_file_path = tk.StringVar()
        # template file path
        self.output_folder_path = tk.StringVar()

        self.master_window = master

        # add title to the window
        master.title("Sessional Diary Maker")

        # make background frame
        self.frame_background = ttk.Frame(master)
        self.frame_background.pack(fill=tk.BOTH, expand=tk.TRUE)

        # make frames
        self.frame_top = ttk.LabelFrame(self.frame_background)
        self.frame_top.pack(padx=10, pady=10, fill=tk.BOTH, expand=tk.TRUE)

        self.frame_top.columnconfigure(0, weight=1)
        self.frame_top.columnconfigure(1, weight=1)

        # Buttons
        input_file_button = ttk.Button(
            self.frame_top,
            text="Input Excel File",
            width=20,
            command=self.get_input_file,
        )
        input_file_button.grid(row=0, column=0, stick="w", padx=10, pady=3)

        template_html_button = ttk.Button(
            self.frame_top,
            text="Output folder",
            width=20,
            command=self.get_output_folder,
        )
        template_html_button.grid(row=1, column=0, stick="w", padx=10, pady=3)

        # checkbox
        self.no_excel = tk.BooleanVar()
        self.no_excel.set(False)
        checkbox = ttk.Checkbutton(
            self.frame_top,
            text="Output Excel file",
            variable=self.no_excel,
            onvalue=False,
            offvalue=True,
        )
        checkbox.grid(row=2, column=0, stick="w", padx=10, pady=3)

        # run button
        run_OP_tool_button = ttk.Button(
            self.frame_top, text="Run", width=12, command=self.gui_run
        )
        run_OP_tool_button.grid(row=7, column=0, columnspan=3, padx=10, pady=10)

    def gui_run(self):

        infilename = self.input_file_path.get()
        output_folder = self.output_folder_path.get()

        # some validation
        infilename_Path = Path(infilename)
        output_folder_Path = Path(output_folder)
        if not (infilename_Path.exists() and infilename.endswith(".xlsx")):
            messagebox.showerror(
                "Error",
                "Please select an Excel file using the Input Excel File button.",
            )
            return
        if not (output_folder_Path.is_dir() and os.access(output_folder, os.W_OK)):
            messagebox.showerror(
                "Error", "Please select a folder for the output files to be saved into"
            )
            return

        self.run_callback(infilename, output_folder, no_excel=self.no_excel.get())
        messagebox.showinfo(title=None, message="All Done!")

    def get_input_file(self):
        directory = filedialog.askopenfilename(
            parent=self.master_window, filetypes=[("Excel files", ".xlsx .xls")]
        )
        self.input_file_path.set(directory)

    def get_output_folder(self):
        directory = ""
        try:
            parent_folder = Path(self.input_file_path.get()).parent
            if parent_folder.is_dir():
                directory = filedialog.askdirectory(
                    parent=self.master_window, initialdir=str(parent_folder.resolve())
                )
            else:
                directory = filedialog.askdirectory(parent=self.master_window)
        except Exception as e:
            print(e)
            directory = filedialog.askdirectory(parent=self.master_window)
        finally:
            self.output_folder_path.set(directory)


def mainloop(run_callback: Callable[[str, str, bool], None]):
    run_OP_toolapp = tk.Tk()
    GUIApp(run_OP_toolapp, run_callback)
    run_OP_toolapp.mainloop()
