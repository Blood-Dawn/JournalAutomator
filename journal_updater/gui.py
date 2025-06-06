import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
from . import journal_updater


def run_gui():
    root = tk.Tk()
    root.title("ABNFF Journal Updater")
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)

    selected_base = tk.StringVar()
    selected_content = tk.StringVar()
    selected_output = tk.StringVar()
    selected_articles: list[str] = []
    articles_label = tk.StringVar()
    volume = tk.StringVar()
    issue = tk.StringVar()
    month_year = tk.StringVar()
    section_title = tk.StringVar()
    cover_page = tk.IntVar(value=1)
    header_page = tk.IntVar(value=2)

    def choose_base():
        path = filedialog.askopenfilename(
            title="Select base DOCX", filetypes=[("Word files", "*.docx")]
        )
        if path:
            selected_base.set(path)

    def choose_content():
        path = filedialog.askdirectory(title="Select content folder")
        if path:
            selected_content.set(path)

    def choose_output():
        path = filedialog.asksaveasfilename(
            title="Save output DOCX",
            defaultextension=".docx",
            filetypes=[("Word files", "*.docx")],
        )
        if path:
            selected_output.set(path)

    def choose_articles():
        paths = filedialog.askopenfilenames(
            title="Select article DOCX files",
            filetypes=[("Word files", "*.docx")],
        )
        if paths:
            selected_articles.clear()
            selected_articles.extend(paths)
            articles_label.set(", ".join(Path(p).name for p in selected_articles))

    def run_update():
        if (
            not selected_base.get()
            or not selected_content.get()
            or not selected_output.get()
            or not volume.get()
            or not issue.get()
            or not month_year.get()
            or not section_title.get()
        ):
            messagebox.showerror(
                "Missing information", "Please select all required paths"
            )
            return
        try:
            output_arg = Path(selected_output.get()) if selected_output.get() else None
            journal_updater.main_from_gui(
                Path(selected_base.get()),
                Path(selected_content.get()),
                Path(selected_output.get()),
                volume.get(),
                issue.get(),
                month_year.get(),
                section_title.get(),
                cover_page.get(),
                header_page.get(),
                [Path(p) for p in selected_articles] if selected_articles else None,
            )
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update journal: {e}")

    frm = tk.Frame(root, padx=10, pady=10)
    frm.pack(fill="both", expand=True)
    frm.columnconfigure(1, weight=1)

    tk.Button(frm, text="Choose Base DOCX", command=choose_base).grid(
        row=0, column=0, sticky="ew"
    )
    tk.Label(frm, textvariable=selected_base, anchor="w").grid(
        row=0, column=1, padx=5, sticky="ew"
    )

    tk.Button(frm, text="Choose Content Folder", command=choose_content).grid(
        row=1, column=0, sticky="ew"
    )
    tk.Label(frm, textvariable=selected_content, anchor="w").grid(
        row=1, column=1, padx=5, sticky="ew"
    )

    tk.Button(frm, text="Choose Output DOCX (optional)", command=choose_output).grid(
        row=2, column=0, sticky="ew"
    )
    tk.Label(frm, textvariable=selected_output, anchor="w").grid(
        row=2, column=1, padx=5, sticky="ew"
    )

    tk.Button(frm, text="Choose Article Files", command=choose_articles).grid(
        row=3, column=0, sticky="ew"
    )
    tk.Label(frm, textvariable=articles_label, anchor="w").grid(
        row=3, column=1, padx=5, sticky="ew"
    )

    row = 4
    tk.Label(frm, text="Volume:").grid(row=row, column=0, sticky="e")
    tk.Entry(frm, textvariable=volume).grid(row=row, column=1, sticky="ew")
    row += 1
    tk.Label(frm, text="Issue:").grid(row=row, column=0, sticky="e")
    tk.Entry(frm, textvariable=issue).grid(row=row, column=1, sticky="ew")
    row += 1
    tk.Label(frm, text="Month/Year:").grid(row=row, column=0, sticky="e")
    tk.Entry(frm, textvariable=month_year).grid(row=row, column=1, sticky="ew")
    row += 1
    tk.Label(frm, text="Section Title:").grid(row=row, column=0, sticky="e")
    tk.Entry(frm, textvariable=section_title).grid(row=row, column=1, sticky="ew")
    row += 1
    tk.Label(frm, text="Cover Page #:").grid(row=row, column=0, sticky="e")
    tk.Entry(frm, textvariable=cover_page).grid(row=row, column=1, sticky="ew")
    row += 1
    tk.Label(frm, text="Header Page #:").grid(row=row, column=0, sticky="e")
    tk.Entry(frm, textvariable=header_page).grid(row=row, column=1, sticky="ew")
    row += 1

    tk.Button(frm, text="Run Update", command=run_update).grid(
        row=row, column=0, columnspan=2, pady=5
    )

    root.mainloop()


if __name__ == "__main__":
    run_gui()
