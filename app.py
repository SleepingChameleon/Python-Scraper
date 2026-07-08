import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES
import pdfplumber
import pandas as pd
import json
import os
import threading
from pathlib import Path

BG        = "#0F1117"
SURFACE   = "#1A1D27"
CARD      = "#20232E"
BORDER    = "#2E3347"
ACCENT    = "#4F8EF7"
ACCENT2   = "#7C5CBF"
SUCCESS   = "#34C77B"
WARNING   = "#F5A623"
ERROR_C   = "#E5534B"
TEXT      = "#E8EAF0"
SUBTEXT   = "#8B90A7"
WHITE     = "#FFFFFF"
 
FONT_TITLE  = ("Segoe UI", 22, "bold")
FONT_HEAD   = ("Segoe UI", 13, "bold")
FONT_BODY   = ("Segoe UI", 10)
FONT_SMALL  = ("Segoe UI", 9)
FONT_MONO   = ("Consolas", 9)
FONT_BTN    = ("Segoe UI", 10, "bold")

def pdf_to_excel(pdf_path: str, out_path: str) -> dict:
    all_tables = []
    page_info = []

    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages, 1):
            tables = page.extract_tables()
            for j, table in enumerate(tables, 1):
                if table and any(any(c for c in row if c) for row in table):
                    header = table[0] if table[0] else [f"Col{k}" for k in range(len(table[0]))]
                    header = [str(h) if h else f"Col{idx}" for idx, h in enumerate(header)]
                    rows = [[str(c) if c is not None else "" for c in row] for row in table[1:]]
                    df = pd.DataFrame(rows, columns=header)
                    all_tables.append(df)
                    page_info.append((i, j))
    
    if not all_tables:
        lines = []
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                txt = page.extract_text() or "" 
                lines.extend(txt.splitlines())
        df = pd.DataFrame({"text": lines})
        all_tables = [df]
        page_info = [(0, 0)]
    
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for idx, (df, (pg, tbl)) in enumerate (zip(all_tables, page_info)):
            sheet = f"P{pg}_T{tbl}" if page_info[0] != (0, 0) else f"Sheet{idx+1}"
            df.to_excel(writer, sheet_name=sheet, index=False)
    
    return {"tables": len(all_tables), "output": out_path}

def excel_to_json(xlsx_path: str, out_path: str) -> dict:
    xl = pd.ExcelFile(xlsx_path)
    result = {}

    for sheet in xl.sheet_names:
        df = xl.parse(sheet)
        df = df.where(pd.notnull(df), None)
        result[sheet] = json.loads(df.to_json(orient="records", date_format="iso"))

    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(result, f, indent=2, ensure_ascii=False)

    total_rows = sum(len(v) for v in result.values())
    return {"sheets": len(result), "rows": total_rows, "output": out_path}

class DropZone(tk.Frame):
    def __init__(self, parent, label, accept_ext, on_drop, icon="📂", **kwargs):
        super().__init__(parent, bg=CARD, **kwargs)
        self.accept_ext = [e.lower() for e in accept_ext]
        self.on_drop = on_drop
        self.file_path = None

        self.configure(highlightbackground=BORDER, highlightthickness=1, relief="flat", cursor="hand2")

        inner = tk.Frame(self, bg=CARD)
        inner.pack(fill="both", expand=True, padx=20, pady=20)

        self.icon_lbl = tk.Label(inner, text=icon, bg=CARD, font=("Segoe UI Emoji", 28), fg=SUBTEXT)
        self.icon_lbl.pack(pady=(0, 6))
        self.title_lbl = tk.Label(inner, text=label, bg=CARD, font=FONT_HEAD, fg=TEXT)
        self.title_lbl.pack()
        self.hint_lbl = tk.Label(inner, text=f"Drop or click  ·  {', '.join(accept_ext).upper()}", bg=CARD, font=FONT_SMALL, fg=SUBTEXT)
        self.hint_lbl.pack(pady=(4, 0))
        self.file_lbl = tk.Label(inner, text="", bg=CARD, font=FONT_SMALL, fg=ACCENT, wraplength=280)
        self.hint_lbl.pack(pady=(8, 0))

        for w in (self, inner, self.icon_lbl, self.title_lbl, self.hint_lbl, self.file_lbl):
            w.bind("<Button-1>", self._browse)
            w.bind("<Enter>", self._hover_on)
            w.bind("<Leave>", self._hover_off)

        self.drop_target_register(DND_FILES)
        self.dnd_bind("<<Drop>>", self._dnd_drop)

    def _hover_on(self, _=None):
        self.configure(highlightbackground=ACCENT, bg=SURFACE)
        self._set_children_bg(self, SURFACE)

    def _hover_off(self, _=None):
        col = SUCCESS if self.file_path else BORDER
        self.configure(highlightbackground=col, bg=CARD)
        self._set_children_bg(self, CARD)

    def _set_children_bg(self, widget, color):
        for child in widget.winfo_children():
            try:
                child.configure(bg=color)
            except tk.TclError:
                pass
            self._set_children_bg(child, color)

    def _dnd_drop(self, event):
        path = event.data.strip().strip("{}")
        self._load_file(path)

    def _browse(self, _=None):
        ext_list = [(f"{e.upper()} files", f"*{e}") for e in self.accept_ext]
        path = filedialog.askopenfilename(filetypes=ext_list + [("All file", "*.*")])
        if path:
            self._load_file(path)

    def _load_file(self, path):
        ext = Path(path).suffix.lower()
        if ext not in self.accept_ext:
            self.configure(highlightbackground=ERROR_C)
            self.file_lbl.configure(text=f"❌  Expected {', '.join(self.accept_ext).upper()}", fg=ERROR_C)
            return
        
        self.file_path = path
        name = Path(path).name
        self.file_lbl.configure(text=f"✔  {name}", fg=SUCCESS)
        self.configure(highlightbackground=SUCCESS)
        self.on_drop(path)

class StatusBar(tk.Frame):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, bg=SURFACE, height=32, **kwargs)
        self.lbl = tk.Label(self, text="Ready", bg=SURFACE, font=FONT_SMALL, fg=SUBTEXT, anchor="w")
        self.lbl.pack(side="left", padx=12, fill="x", expand=True)
        self.prog = ttk.Progressbar(self, mode="indeterminate", length=100, style="Accent.Horizontal.TProgressbar")
    
    def set(self, msg, color=SUBTEXT, working=False):
        self.lbl.configure(text=msg, fg=color)
        if working:
            self.prog.pack(side="right", padx=12, pady=6)
            self.prog.start(12)
        else:
            self.prog.stop()
            self.prog.pack_forget()

class LogBox(tk.Frame):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, bg=CARD, **kwargs)
        self.configure(highlightbackground=BORDER, highlightthickness=1)

        tk.Label(self, text=" 📋  Activity Log", bg=CARD, font=FONT_SMALL, fg=SUBTEXT, anchor='w').pack(fill='x', pady=(6, 0))
        self.text = tk.Text(self, bg=BG, fg=TEXT, font=FONT_MONO, relief='flat', bd=0, state="disabled", wrap="word", height=7)
        scroll = tk.Scrollbar(self, command=self.text.yview, bg=CARD)
        self.text.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.text.pack(fill="both", expand=True, padx=8, pady=4)

        self.text.tag_configure("ok", foreground=SUCCESS)
        self.text.tag_configure("err", foreground=ERROR_C)
        self.text.tag_configure("info", foreground=ACCENT)
        self.text.tag_configure("warn", foreground=WARNING)
    
    def log(self, msg, tag="info"):
        self.text.configure(state="normal")
        self.text.insert("end", msg + "\n", tag)
        self.text.see("end")
        self.text.configure(state="disabled")

class PdfToExcelPanel(tk.Frame):
    def __init__(self, parent, log, status, **kwargs):
        super().__init__(parent, bg=BG, **kwargs)
        self.log = log
        self.status = status
        self.source = None
        self._out_path = None
        self._build()

    def _build(self):
        hdr = tk.Frame(self, bg=BG)
        hdr.pack(fill="x", padx=24, pady=(24, 0))
        tk.Label(hdr, text="PDF  →  Excel", bg=BG, font=FONT_TITLE, fg=WHITE).pack(side="left")
        tk.Label(hdr, text="  Table Extractor  ", bg=ACCENT2, fg=WHITE, font=FONT_SMALL, padx=6, pady=2).pack(side="left", padx=(12, 0), anchor="s")
        tk.Label(self, text="Drop a PDF and extract its tables into an Excel workbook.", bg=BG, font=FONT_BODY, fg=SUBTEXT).pack(anchor="w", padx=24, pady=(4, 16))
        self.drop = DropZone(self, "Drop PDF Here", [".pdf"], on_drop=self._file_chosen, icon="📄")
        self.drop.pack(fill="x", padx=24, ipady=20)
        self.btn = tk.Button(self, text="⚡  Convert to Excel", bg=ACCENT, fg=WHITE, font=FONT_BTN, relief="flat", bd=0, padx=24, pady=10, activebackground="#3a72d4", activeforeground=WHITE, cursor="hand2", command=self._convert)
        self.btn.pack(pady=16, padx=24, anchor="w")
        self.result_frame = tk.Frame(self, bg=CARD, highlightbackground=SUCCESS, highlightthickness=0)
        self.result_lbl = tk.Label(self.result_frame, text="", bg=CARD, font=FONT_BODY, fg=SUCCESS, wraplength=500, justify="left")
        self.open_btn = tk.Button(self.result_frame, text="📂  Open folder", bg=SUCCESS, fg=WHITE, font=FONT_BTN, relief="flat", padx=12, pady=6, cursor="hand2", command=self._open_folder)
    
    def _file_chosen(self, path):
        self.source = path
        self.log.log(f"→  PDF selected: {Path(path).name}", "info")
        self.status.set(f"PDF ready: {Path(path).name}")

    def _convert(self):
        if not self.source:
            messagebox.showwarning("No file", "Please drop or select a PDF first.")
            return
        out_path = str(Path(self.source).parent / (Path(self.source).stem + "_converted.xlsx"))
        self._out_path = out_path
        self.btn.configure(state="disabled", text="Converting...")
        self.status.set("Extracting tables from PDF…", ACCENT, working=True)
        self.log.log(f"⏳  Converting: {Path(self.source).name}", "info")

        def task():
            try:
                result = pdf_to_excel(self.source, out_path)
                self.after(0, self._on_success, result)
            except Exception as e:
                self.after(0, self._on_error, str(e))
        
        threading.Thread(target=task, daemon=True).start()
    
    def _on_success(self, result):
        self.btn.configure(state="normal", text="⚡  Convert to Excel")
        self.status.set("Conversion complete!", SUCCESS)
        self.result_lbl.configure(text=f"✔  Done!  {result['tables']} sheet(s) created  ·  {Path(result['output']).name}")
        self.result_frame.configure(highlightthickness=1)
        self.result_frame.pack(fill="x", padx=24, pady=(0, 12))
        self.result_lbl.pack(anchor="w", padx=12, pady=8)
        self.open_btn.pack(anchor="e", padx=12, pady=(0, 8))
        self.log.log(f"✔  Saved: {result['output']}", "ok")
    
    def _on_error(self, msg):
        self.btn.configure(state="normal", text="⚡  Convert to Excel")
        self.log.log(f"✘  Error: {msg}", "err")
        self.status.set(f"Failed: {msg}", ERROR_C)
    
    def _open_folder(self):
        if self._out_path:
            folder = str(Path(self._out_path).parent)
            if os.name == "nt":
                os.startfile(folder)
            else:
                os.system(f'xdg-open "{folder}"')

class ExcelToJsonPanel(tk.Frame):
    def __init__(self, parent, log, status, **kwargs):
        super().__init__(parent, bg=BG, **kwargs)
        self.log = log
        self.status = status
        self.source = None
        self._json_str = ""
        self._build()

    def _build(self):
        hdr = tk.Frame(self, bg=BG)
        hdr.pack(fill="x", padx=24, pady=(24, 0))
        tk.Label(hdr, text="Excel  →  JSON", bg=BG, font=FONT_TITLE, fg=WHITE).pack(side="left")
        tk.Label(hdr, text="  API Body Builder  ", bg=ACCENT, fg=WHITE, font=FONT_SMALL, padx=6, pady=2).pack(side="left", padx=(12, 0), anchor="s")
        tk.Label(self, text="Drop an Excel file and get a structured JSON body for your API.", bg=BG, font=FONT_BODY, fg=SUBTEXT).pack(anchor="w", padx=24, pady=(4, 16))
        self.drop = DropZone(self, "Drop Excel Here", [".xlsx", ".xls"], on_drop=self._file_chosen, icon="📊")
        self.drop.pack(fill='x', padx=24, ipady=20)

        opt = tk.Frame(self, bg=BG)
        opt.pack(fill='x', padx=24, pady=(12, 0))
        tk.Label(opt, text="Preview rows: ", bg=BG, font=FONT_BODY, fg=SUBTEXT).pack(side="left")
        self.preview_var = tk.IntVar(value=10)
        tk.Spinbox(opt, from_=1, to=100, textvariable=self.preview_var, width=5, bg=CARD, fg=TEXT, font=FONT_BODY, buttonbackground=SURFACE, relief="flat").pack(side="left", padx=6)
        tk.Label(opt, text="(preview only — full file always saved)", bg=BG, font=FONT_SMALL, fg=SUBTEXT).pack(side="left", padx=(4, 0))
        self.btn = tk.Button(self, text="⚡  Convert to JSON", bg=ACCENT2, fg=WHITE, font=FONT_BTN, relief="flat", bd=0, padx=24, pady=10, activebackground="#6248a0", activeforeground=WHITE, cursor="hand2", command=self._convert)
        self.btn.pack(pady=16, padx=24, anchor="w")

        ph = tk.Frame(self, bg=BG)
        ph.pack(fill="x", padx=24)
        tk.Label(ph, text="JSON Preview", bg=BG, font=("Segoe UI", 10, "bold"), fg=SUBTEXT).pack(side="left")
        self.copy_btn = tk.Button(ph, text="Copy JSON", bg=SURFACE, fg=SUBTEXT, font=FONT_SMALL, relief="flat", padx=8, pady=3, cursor="hand2", command=self._copy_json)
        self.copy_btn.pack(side="right")
        wrap = tk.Frame(self, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
        wrap.pack(fill="both", expand=True, padx=24, pady=(6, 0))
        self.preview = tk.Text(wrap, bg=BG, fg=TEXT, font=FONT_MONO, relief="flat", bd=0, state="disabled", wrap="none", height=12)
        ys = tk.Scrollbar(wrap, command=self.preview.yview)
        xs = tk.Scrollbar(wrap, orient="horizontal", command=self.preview.xview)
        self.preview.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        ys.pack(side="right", fill='y')
        xs.pack(side='bottom', fill='x')
        self.preview.pack(fill='both', expand=True, padx=4, pady=4)

        self.preview.tag_configure('key', foreground=ACCENT)
        self.preview.tag_configure('str', foreground=SUCCESS)
        self.preview.tag_configure('num', foreground=WARNING)
        self.preview.tag_configure('null', foreground=ERROR_C)
        self.preview.tag_configure('brace', foreground=SUBTEXT)
    
    def _file_chosen(self, path):
        self.source = path
        self.log.log(f"→  Excel selected: {Path(path).name}", "info")
        self.status.set(f"Excel ready: {Path(path).name}")

    def _convert(self):
        if not self.source:
            messagebox.showwarning("No file", "Please drop or select an Excel file first.")
            return
        out_path = str(Path(self.source).parent / (Path(self.source).stem + "_body.json"))
        self._out_path = out_path
        self.btn.configure(state='disabled', text='Converting...')
        self.status.set("Parsing Excel sheets…", ACCENT2, working=True)
        self.log.log(f"⏳  Converting: {Path(self.source).name}", "info")
    
        def task():
            try:
                result = excel_to_json(self.source, out_path)
                self.after(0, self._on_success, result)
            except Exception as e:
                self.after(0, self._on_error, str(e))
        
        threading.Thread(target=task, daemon=True).start()

    def _on_success(self, result):
        self.btn.configure(state='normal', text="⚡  Convert to JSON")
        self.status.set(f"Done!  {result['sheets']} sheet(s)  ·  {result['rows']} rows  →  {Path(result['output']).name}", SUCCESS)
        self.log.log(f"✔  JSON saved: {result['output']}  |  " f"{result['sheets']} sheet(s), {result['rows']} rows", "ok")
        
        with open(result['output'], encoding='utf-8') as f:
            full = json.load(f)
        
        n = self.preview_var.get()
        preview = {sheet: rows[:n] for sheet, rows in full.items()}
        self._json_str = json.dumps(full, indent=2, ensure_ascii=False)
        self._render_json(json.dumps(preview, indent=2, ensure_ascii=False))
    
    def _on_error(self, msg):
        self.btn.configure(state='normal', text='⚡  Convert to JSON')
        self.log.log(f"✘  Error: {msg}", "err")
        self.status.set(f"Failed: {msg}", ERROR_C)
    
    def _render_json(self, text):
        import re
        self.preview.configure(state='normal')
        self.preview.delete("1.0", "end")

        token_re = re.compile(r'"((?:[^"\\]|\\.)*)"\s*:'
                               r'|"((?:[^"\\]|\\.)*)"'
                               r'|(-?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?)'
                               r'|(null|true|false)'
                               r'|([{}\[\],])')
        pos = 0
        for m in token_re.finditer(text):
            if m.start() > pos:
                self.preview.insert("end", text[pos:m.start()])
            if m.group(1) is not None: self.preview.insert("end", f'"{m.group(1)}":', "key")
            elif m.group(2) is not None: self.preview.insert("end", f'"{m.group(2)}"', "str")
            elif m.group(3) is not None: self.preview.insert("end", m.group(3), "num")
            elif m.group(4) is not None: self.preview.insert("end", m.group(4), "null")
            elif m.group(5) is not None: self.preview.insert("end", m.group(5), "brace")
            pos = m.end()
        if pos < len(text):
            self.preview.insert("end", text[pos:])
        self.preview.configure(state='disabled')
    
    def _copy_json(self):
        if self._json_str:
            self.clipboard_clear()
            self.clipboard_append(self._json_str)
            self.copy_btn.configure(text="Copied!", fg=SUCCESS)
            self.after(2000, lambda: self.copy_btn.configure(text="COPY JSON", fg=SUBTEXT))

class App(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("ProductScraper  ·  PDF & Excel Tools")
        self.configure(bg=BG)
        self.geometry("780x780")
        self.minsize(680, 600)
        self._center()
        self._apply_style()
        self._build()

    def _center(self):
        self.update_idletasks()
        x = (self.winfo_screenwidth() - 780) // 2
        y = (self.winfo_screenheight() - 720) // 2
        self.geometry(f"780x720+{x}+{y}")
    
    def _apply_style(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure("TNotebook", background=BG, borderwidth=0)
        s.configure("TNotebook.Tab", background=SURFACE, foreground=SUBTEXT, font=FONT_BTN, padding=[18, 8], borderwidth=0)
        s.map("TNotebook.Tab", background=[("selected", CARD), ("active", CARD)], foreground=[("selected", WHITE), ("active", TEXT)])
        s.configure("Accent.Horizontal.TProgressbar", troughcolor=SURFACE, background=ACCENT, borderwidth=0)
    
    def _build(self):
        top = tk.Frame(self, bg=SURFACE, height=52)
        top.pack(fill='x')
        top.pack_propagate(False)
        tk.Label(top, text="  ⚡  ConverterKit", bg=SURFACE, font=("Segoe UI", 14, "bold"), fg=WHITE).pack(side="left", padx=8)
        tk.Label(top, text="PDF & Excel Utilities", bg=SURFACE, font=FONT_SMALL, fg=SUBTEXT).pack(side="left")

        bottom = tk.Frame(self, bg=BG)
        bottom.pack(fill='x', side='bottom')

        self.status = StatusBar(bottom)
        self.status.pack(fill='x', side='bottom')

        self.log = LogBox(bottom)
        self.log.pack(fill='x')

        self.nb = ttk.Notebook(self)
        self.nb.pack(fill='both', expand=True)

        t1 = PdfToExcelPanel(self.nb, self.log, self.status)
        t2 = ExcelToJsonPanel(self.nb, self.log, self.status)
        self.nb.add(t1, text="  📄  PDF → Excel  ")
        self.nb.add(t2, text="  📊  Excel → JSON  ")

        self.log.log("ProductPriceScraper ready. Drop a file to get started.", "info")

if __name__ == "__main__":
    App().mainloop()