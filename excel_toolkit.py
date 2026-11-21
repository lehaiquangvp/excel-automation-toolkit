"""
Excel Automation Toolkit - desktop app (no coding needed).

Run:  python excel_toolkit.py
"""

import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import excel_core as core

APP_TITLE = "Excel Automation Toolkit"
FILETYPES = [("Excel / CSV files", "*.xlsx *.xlsm *.xls *.csv"), ("All files", "*.*")]


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("720x560")
        self.minsize(640, 500)

        style = ttk.Style(self)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        style.configure("Run.TButton", font=("Segoe UI", 10, "bold"), padding=6)

        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True, padx=10, pady=(10, 0))
        notebook.add(self._merge_tab(notebook), text="  1. Merge files  ")
        notebook.add(self._split_tab(notebook), text="  2. Split  ")
        notebook.add(self._dedupe_tab(notebook), text="  3. Remove duplicates  ")

        log_frame = ttk.LabelFrame(self, text="Result")
        log_frame.pack(fill="both", padx=10, pady=10)
        self.log = tk.Text(log_frame, height=8, wrap="word", state="disabled",
                           font=("Consolas", 9))
        self.log.pack(fill="both", expand=True, padx=5, pady=5)

    # ------------------------------------------------------------ shared UI
    def _write_log(self, text):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.insert("end", text)
        self.log.configure(state="disabled")

    def _run(self, button, prepare):
        """
        prepare() runs here (UI thread): it reads the form and returns `job`.
        job() then runs in the background so the window does not freeze.
        """
        try:
            job = prepare()
        except Exception as e:
            self._done(button, f"Error: {e}", False)
            return
        button.configure(state="disabled")
        self._write_log("Working... please wait.")
        self.configure(cursor="watch")

        result = queue.Queue()

        def worker():  # never touch Tk widgets from this thread
            try:
                result.put((job(), True))
            except PermissionError:
                result.put(("Could not save the file. If it is open in Excel, "
                            "close it and try again.", False))
            except Exception as e:  # show any problem in plain words
                result.put((f"Error: {e}", False))

        def poll():
            try:
                msg, ok = result.get_nowait()
            except queue.Empty:
                self.after(100, poll)
                return
            self._done(button, msg, ok)

        threading.Thread(target=worker, daemon=True).start()
        self.after(100, poll)

    def _done(self, button, msg, ok):
        button.configure(state="normal")
        self.configure(cursor="")
        self._write_log(msg)
        if ok:
            messagebox.showinfo(APP_TITLE, "Done!\n\n" + msg.splitlines()[0])
        else:
            messagebox.showerror(APP_TITLE, msg)

    @staticmethod
    def _row(parent, label, var, browse):
        row = ttk.Frame(parent)
        row.pack(fill="x", padx=10, pady=4)
        ttk.Label(row, text=label, width=16).pack(side="left")
        ttk.Entry(row, textvariable=var).pack(side="left", fill="x", expand=True, padx=4)
        ttk.Button(row, text="Browse...", command=browse).pack(side="left")

    # ------------------------------------------------------------ 1. merge
    def _merge_tab(self, parent):
        tab = ttk.Frame(parent)
        ttk.Label(tab, text="Combine many Excel/CSV files into one workbook.",
                  font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=10, pady=(10, 4))

        list_frame = ttk.Frame(tab)
        list_frame.pack(fill="both", expand=True, padx=10)
        self.merge_list = tk.Listbox(list_frame, height=7, selectmode="extended")
        self.merge_list.pack(side="left", fill="both", expand=True)
        btns = ttk.Frame(list_frame)
        btns.pack(side="left", fill="y", padx=(6, 0))
        ttk.Button(btns, text="Add files...", command=self._merge_add_files).pack(fill="x")
        ttk.Button(btns, text="Add folder...", command=self._merge_add_folder).pack(fill="x", pady=4)
        ttk.Button(btns, text="Remove selected", command=self._merge_remove).pack(fill="x")
        ttk.Button(btns, text="Clear all",
                   command=lambda: self.merge_list.delete(0, "end")).pack(fill="x", pady=4)

        self.merge_mode = tk.StringVar(value="stack")
        opts = ttk.Frame(tab)
        opts.pack(fill="x", padx=10, pady=6)
        ttk.Radiobutton(opts, text="Stack all rows into ONE sheet (columns matched by header)",
                        variable=self.merge_mode, value="stack").pack(anchor="w")
        ttk.Radiobutton(opts, text="Keep each sheet as a separate sheet",
                        variable=self.merge_mode, value="sheets").pack(anchor="w")
        self.merge_source = tk.BooleanVar(value=True)
        ttk.Checkbutton(opts, text='Add "Source File" and "Source Sheet" columns',
                        variable=self.merge_source).pack(anchor="w", pady=(4, 0))

        self.merge_out = tk.StringVar(value=str(Path.home() / "Desktop" / "Merged.xlsx"))
        self._row(tab, "Save as:", self.merge_out, lambda: self._save_as(self.merge_out))

        btn = ttk.Button(tab, text="Merge files", style="Run.TButton")
        btn.configure(command=lambda: self._run(btn, self._merge_job))
        btn.pack(pady=8)
        return tab

    def _merge_job(self):
        files = list(self.merge_list.get(0, "end"))
        if not files:
            raise ValueError("Please add some files first.")
        out, mode, src = self.merge_out.get(), self.merge_mode.get(), self.merge_source.get()
        return lambda: core.merge_files(files, out, mode, src)

    def _merge_add_files(self):
        for f in filedialog.askopenfilenames(filetypes=FILETYPES):
            self.merge_list.insert("end", f)

    def _merge_add_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            files = core.collect_files([folder])
            if not files:
                messagebox.showwarning(APP_TITLE, "No Excel or CSV files in that folder.")
            for f in files:
                self.merge_list.insert("end", str(f))

    def _merge_remove(self):
        for i in reversed(self.merge_list.curselection()):
            self.merge_list.delete(i)

    # ------------------------------------------------------------ 2. split
    def _split_tab(self, parent):
        tab = ttk.Frame(parent)
        ttk.Label(tab, text="Split one workbook into several files.",
                  font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=10, pady=(10, 4))

        self.split_in = tk.StringVar()
        self._row(tab, "Workbook:", self.split_in, self._split_pick)

        self.split_mode = tk.StringVar(value="sheets")
        box = ttk.LabelFrame(tab, text="How to split")
        box.pack(fill="x", padx=10, pady=6)
        ttk.Radiobutton(box, text="One file per SHEET", variable=self.split_mode,
                        value="sheets", command=self._split_toggle).pack(anchor="w", padx=6)
        ttk.Radiobutton(box, text="By the values of a COLUMN (e.g. one file per Region)",
                        variable=self.split_mode, value="column",
                        command=self._split_toggle).pack(anchor="w", padx=6)

        self.split_col_frame = ttk.Frame(box)
        self.split_col_frame.pack(fill="x", padx=24, pady=4)
        ttk.Label(self.split_col_frame, text="Sheet:").grid(row=0, column=0, sticky="w")
        self.split_sheet = ttk.Combobox(self.split_col_frame, state="readonly", width=28)
        self.split_sheet.grid(row=0, column=1, padx=4, pady=2, sticky="w")
        self.split_sheet.bind("<<ComboboxSelected>>", lambda e: self._split_load_cols())
        ttk.Label(self.split_col_frame, text="Column:").grid(row=1, column=0, sticky="w")
        self.split_col = ttk.Combobox(self.split_col_frame, state="readonly", width=28)
        self.split_col.grid(row=1, column=1, padx=4, pady=2, sticky="w")
        self.split_as_sheets = tk.BooleanVar(value=False)
        ttk.Checkbutton(self.split_col_frame, text="Put results in ONE file (one sheet per value)",
                        variable=self.split_as_sheets).grid(row=2, column=0, columnspan=2,
                                                            sticky="w", pady=2)

        self.split_out = tk.StringVar(value=str(Path.home() / "Desktop" / "Split files"))
        self._row(tab, "Output folder:", self.split_out, self._split_pick_dir)

        btn = ttk.Button(tab, text="Split", style="Run.TButton")
        btn.configure(command=lambda: self._run(btn, self._split_job))
        btn.pack(pady=8)
        self._split_toggle()
        return tab

    def _split_job(self):
        src, out = self.split_in.get(), self.split_out.get()
        if not src:
            raise ValueError("Please choose a workbook first.")
        if self.split_mode.get() == "sheets":
            return lambda: core.split_sheets(src, out)
        col, sheet, as_sheets = (self.split_col.get(), self.split_sheet.get() or None,
                                 self.split_as_sheets.get())
        if not col:
            raise ValueError("Please choose a column to split by.")
        return lambda: core.split_by_column(src, col, out, sheet, as_sheets)

    def _split_toggle(self):
        state = "readonly" if self.split_mode.get() == "column" else "disabled"
        self.split_sheet.configure(state=state)
        self.split_col.configure(state=state)

    def _split_pick(self):
        f = filedialog.askopenfilename(filetypes=FILETYPES)
        if not f:
            return
        self.split_in.set(f)
        try:
            sheets = core.list_sheets(f)
            self.split_sheet["values"] = sheets
            self.split_sheet.set(sheets[0] if sheets else "")
            self._split_load_cols()
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Could not read file: {e}")

    def _split_load_cols(self):
        try:
            cols = core.list_columns(self.split_in.get(), self.split_sheet.get() or None)
            self.split_col["values"] = cols
            self.split_col.set("")
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Could not read columns: {e}")

    def _split_pick_dir(self):
        d = filedialog.askdirectory()
        if d:
            self.split_out.set(d)

    # ------------------------------------------------------------ 3. dedupe
    def _dedupe_tab(self, parent):
        tab = ttk.Frame(parent)
        ttk.Label(tab, text="Remove duplicate rows from every sheet.",
                  font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=10, pady=(10, 4))

        self.dd_in = tk.StringVar()
        self._row(tab, "Workbook:", self.dd_in, self._dd_pick)

        box = ttk.LabelFrame(tab, text="Compare rows using these columns "
                                       "(select none = whole row must match)")
        box.pack(fill="both", expand=True, padx=10, pady=6)
        self.dd_cols = tk.Listbox(box, selectmode="multiple", height=6, exportselection=False)
        self.dd_cols.pack(fill="both", expand=True, padx=6, pady=6)

        opts = ttk.Frame(tab)
        opts.pack(fill="x", padx=10)
        self.dd_keep = tk.StringVar(value="first")
        ttk.Label(opts, text="Keep:").pack(side="left")
        ttk.Radiobutton(opts, text="first copy", variable=self.dd_keep,
                        value="first").pack(side="left", padx=4)
        ttk.Radiobutton(opts, text="last copy", variable=self.dd_keep,
                        value="last").pack(side="left", padx=4)
        self.dd_loose = tk.BooleanVar(value=True)
        ttk.Checkbutton(opts, text="Ignore upper/lower case and extra spaces",
                        variable=self.dd_loose).pack(side="left", padx=16)

        self.dd_out = tk.StringVar(value=str(Path.home() / "Desktop" / "Cleaned.xlsx"))
        self._row(tab, "Save as:", self.dd_out, lambda: self._save_as(self.dd_out))
        ttk.Label(tab, text='Removed rows are kept in a "Removed Duplicates" sheet, '
                            "and your original file is never changed.",
                  foreground="gray").pack(anchor="w", padx=10)

        btn = ttk.Button(tab, text="Remove duplicates", style="Run.TButton")
        btn.configure(command=lambda: self._run(btn, self._dd_job))
        btn.pack(pady=8)
        return tab

    def _dd_job(self):
        src, out = self.dd_in.get(), self.dd_out.get()
        if not src:
            raise ValueError("Please choose a workbook first.")
        cols = [self.dd_cols.get(i) for i in self.dd_cols.curselection()] or None
        keep, loose = self.dd_keep.get(), self.dd_loose.get()
        return lambda: core.remove_duplicates(src, out, cols, keep, loose)

    def _dd_pick(self):
        f = filedialog.askopenfilename(filetypes=FILETYPES)
        if not f:
            return
        self.dd_in.set(f)
        self.dd_cols.delete(0, "end")
        try:
            for c in core.list_columns(f):
                self.dd_cols.insert("end", c)
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Could not read file: {e}")

    # ------------------------------------------------------------ misc
    @staticmethod
    def _save_as(var):
        f = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                         filetypes=[("Excel workbook", "*.xlsx")],
                                         initialfile=Path(var.get()).name)
        if f:
            var.set(f)


if __name__ == "__main__":
    App().mainloop()
