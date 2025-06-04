import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
from . import journal_updater


def run_gui():
    root = tk.Tk()
    root.title("ABNFF Journal Updater")

    selected_base = tk.StringVar()
    selected_content = tk.StringVar()
    selected_output = tk.StringVar()

    def choose_base():
        path = filedialog.askopenfilename(title="Select base DOCX", filetypes=[("Word files", "*.docx")])
        if path:
            selected_base.set(path)

    def choose_content():
        path = filedialog.askdirectory(title="Select content folder")
        if path:
            selected_content.set(path)

    def choose_output():
        path = filedialog.asksaveasfilename(title="Save output DOCX", defaultextension=".docx", filetypes=[("Word files", "*.docx")])
        if path:
            selected_output.set(path)

    def run_update():
        if not selected_base.get() or not selected_content.get() or not selected_output.get():
            messagebox.showerror("Missing information", "Please select all required paths")
            return
        try:
            journal_updater.main_from_gui(
                Path(selected_base.get()),
                Path(selected_content.get()),
                Path(selected_output.get()),
            )
            messagebox.showinfo("Success", "Journal updated successfully")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update journal: {e}")

    frm = tk.Frame(root, padx=10, pady=10)
    frm.pack()

    tk.Button(frm, text="Choose Base DOCX", command=choose_base).grid(row=0, column=0, sticky="ew")
    tk.Label(frm, textvariable=selected_base, width=40, anchor="w").grid(row=0, column=1, padx=5)

    tk.Button(frm, text="Choose Content Folder", command=choose_content).grid(row=1, column=0, sticky="ew")
    tk.Label(frm, textvariable=selected_content, width=40, anchor="w").grid(row=1, column=1, padx=5)

    tk.Button(frm, text="Choose Output DOCX", command=choose_output).grid(row=2, column=0, sticky="ew")
    tk.Label(frm, textvariable=selected_output, width=40, anchor="w").grid(row=2, column=1, padx=5)

    tk.Button(frm, text="Run Update", command=run_update).grid(row=3, column=0, columnspan=2, pady=5)

    root.mainloop()


if __name__ == "__main__":
    run_gui()
