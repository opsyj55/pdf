"""
PDF Fit & Print - resize every page of a PDF to a custom paper size, then save or print.
Requires: Python 3.8+ and pypdf   ->   pip install pypdf
Run:      python pdf_fit_print.py
Build:    pyinstaller --onefile --windowed pdf_fit_print.py
"""
import os
import sys
import subprocess
import tempfile
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from pypdf import PdfReader, PdfWriter, PageObject, Transformation

PT_PER = {"mm": 72 / 25.4, "cm": 72 / 2.54, "in": 72.0, "pt": 1.0}

# name: (width, height, unit) in portrait
PRESETS = {
    "Custom": None,
    "A3": (297, 420, "mm"),
    "A4": (210, 297, "mm"),
    "A5": (148, 210, "mm"),
    "A6": (105, 148, "mm"),
    "B5": (176, 250, "mm"),
    "Letter": (8.5, 11, "in"),
    "Legal": (8.5, 14, "in"),
    "Tabloid": (11, 17, "in"),
    "4 x 6 in (photo)": (4, 6, "in"),
    "ID card (85.6 x 54 mm)": (54, 85.6, "mm"),
}


def resize_pdf(src, dst, width_pt, height_pt, margin_pt=0.0, mode="fit",
               auto_orient=True, progress=None):
    """Write a copy of `src` to `dst` with every page placed on a width x height sheet.

    mode = "fit"     -> keep proportions, centre on sheet, respect margin
    mode = "stretch" -> stretch to fill the area inside the margin
    """
    reader = PdfReader(src)
    if reader.is_encrypted:
        reader.decrypt("")
    writer = PdfWriter()
    total = len(reader.pages)

    for idx, page in enumerate(reader.pages, 1):
        page.transfer_rotation_to_content()  # bake /Rotate into the content
        mb = page.mediabox
        w, h = float(mb.width), float(mb.height)
        left, bottom = float(mb.left), float(mb.bottom)

        W, H = width_pt, height_pt
        if auto_orient and ((w > h) != (W > H)):
            W, H = H, W  # match the page's own orientation

        aw, ah = max(W - 2 * margin_pt, 1), max(H - 2 * margin_pt, 1)
        if mode == "stretch":
            sx, sy = aw / w, ah / h
            tx, ty = margin_pt, margin_pt
        else:
            s = min(aw / w, ah / h)
            sx = sy = s
            tx = (W - w * s) / 2
            ty = (H - h * s) / 2

        sheet = PageObject.create_blank_page(width=W, height=H)
        op = (Transformation().translate(-left, -bottom)
              .scale(sx, sy).translate(tx, ty))
        sheet.merge_transformed_page(page, op)
        writer.add_page(sheet)
        if progress:
            progress(idx, total)

    with open(dst, "wb") as f:
        writer.write(f)
    return total


def send_to_printer(path):
    if sys.platform.startswith("win"):
        os.startfile(path, "print")  # default PDF app, default printer
    elif sys.platform == "darwin":
        subprocess.check_call(["lp", path])
    else:
        subprocess.check_call(["lp", path])


def open_file(path):
    if sys.platform.startswith("win"):
        os.startfile(path)
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF Fit & Print")
        self.geometry("560x520")
        self.resizable(False, False)

        self.src = tk.StringVar()
        self.preset = tk.StringVar(value="A4")
        self.width = tk.StringVar(value="210")
        self.height = tk.StringVar(value="297")
        self.unit = tk.StringVar(value="mm")
        self.margin = tk.StringVar(value="0")
        self.mode = tk.StringVar(value="fit")
        self.auto_orient = tk.BooleanVar(value=True)
        self.busy = False
        self._build()

    def _build(self):
        p = {"padx": 10, "pady": 5}
        f = ttk.Frame(self)
        f.pack(fill="both", expand=True, padx=8, pady=8)

        ttk.Label(f, text="PDF file:").grid(row=0, column=0, sticky="w", **p)
        ttk.Entry(f, textvariable=self.src, width=40).grid(row=0, column=1, columnspan=2, sticky="ew", **p)
        ttk.Button(f, text="Browse…", command=self.pick).grid(row=0, column=3, **p)

        box = ttk.LabelFrame(f, text="Target paper size")
        box.grid(row=1, column=0, columnspan=4, sticky="ew", **p)
        ttk.Label(box, text="Preset:").grid(row=0, column=0, sticky="w", **p)
        cb = ttk.Combobox(box, textvariable=self.preset, values=list(PRESETS),
                          state="readonly", width=24)
        cb.grid(row=0, column=1, columnspan=3, sticky="w", **p)
        cb.bind("<<ComboboxSelected>>", self.apply_preset)

        ttk.Label(box, text="Width:").grid(row=1, column=0, sticky="w", **p)
        ttk.Entry(box, textvariable=self.width, width=9).grid(row=1, column=1, **p)
        ttk.Label(box, text="Height:").grid(row=1, column=2, sticky="e", **p)
        ttk.Entry(box, textvariable=self.height, width=9).grid(row=1, column=3, **p)
        ttk.Label(box, text="Unit:").grid(row=2, column=0, sticky="w", **p)
        ttk.Combobox(box, textvariable=self.unit, values=list(PT_PER),
                     state="readonly", width=6).grid(row=2, column=1, **p)
        ttk.Label(box, text="Margin:").grid(row=2, column=2, sticky="e", **p)
        ttk.Entry(box, textvariable=self.margin, width=9).grid(row=2, column=3, **p)

        opt = ttk.LabelFrame(f, text="How to fit")
        opt.grid(row=2, column=0, columnspan=4, sticky="ew", **p)
        ttk.Radiobutton(opt, text="Fit inside (keep proportions, centred)",
                        variable=self.mode, value="fit").pack(anchor="w", padx=10, pady=2)
        ttk.Radiobutton(opt, text="Stretch to fill (may distort)",
                        variable=self.mode, value="stretch").pack(anchor="w", padx=10, pady=2)
        ttk.Checkbutton(opt, text="Auto-rotate sheet to match each page's orientation",
                        variable=self.auto_orient).pack(anchor="w", padx=10, pady=2)

        self.progress = ttk.Progressbar(f, mode="determinate")
        self.progress.grid(row=3, column=0, columnspan=4, sticky="ew", **p)
        self.status = ttk.Label(f, text="Choose a PDF to begin.")
        self.status.grid(row=4, column=0, columnspan=4, sticky="w", **p)

        bf = ttk.Frame(f)
        bf.grid(row=5, column=0, columnspan=4, pady=10)
        self.btn_save = ttk.Button(bf, text="Save resized PDF…", command=lambda: self.run("save"))
        self.btn_print = ttk.Button(bf, text="Print", command=lambda: self.run("print"))
        self.btn_preview = ttk.Button(bf, text="Preview", command=lambda: self.run("preview"))
        for b in (self.btn_save, self.btn_preview, self.btn_print):
            b.pack(side="left", padx=6)
        f.columnconfigure(1, weight=1)

    def pick(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path:
            self.src.set(path)
            try:
                n = len(PdfReader(path).pages)
                self.status.config(text=f"{n} page(s) loaded.")
            except Exception as e:
                messagebox.showerror("Error", f"Cannot read PDF:\n{e}")

    def apply_preset(self, _=None):
        data = PRESETS.get(self.preset.get())
        if data:
            self.width.set(str(data[0]))
            self.height.set(str(data[1]))
            self.unit.set(data[2])

    def read_size(self):
        try:
            k = PT_PER[self.unit.get()]
            w, h = float(self.width.get()) * k, float(self.height.get()) * k
            m = float(self.margin.get() or 0) * k
        except ValueError:
            raise ValueError("Width, height and margin must be numbers.")
        if w <= 0 or h <= 0:
            raise ValueError("Width and height must be greater than zero.")
        if 2 * m >= min(w, h):
            raise ValueError("Margin is too large for this paper size.")
        return w, h, m

    def run(self, action):
        if self.busy:
            return
        src = self.src.get().strip()
        if not os.path.isfile(src):
            messagebox.showwarning("PDF Fit & Print", "Please select a PDF file first.")
            return
        try:
            w, h, m = self.read_size()
        except ValueError as e:
            messagebox.showwarning("PDF Fit & Print", str(e))
            return

        if action == "save":
            base = os.path.splitext(os.path.basename(src))[0]
            dst = filedialog.asksaveasfilename(defaultextension=".pdf",
                                               initialfile=f"{base}_resized.pdf",
                                               filetypes=[("PDF files", "*.pdf")])
            if not dst:
                return
        else:
            dst = os.path.join(tempfile.gettempdir(), "pdf_fit_print_temp.pdf")

        self.busy = True
        for b in (self.btn_save, self.btn_print, self.btn_preview):
            b.config(state="disabled")
        threading.Thread(target=self.work, args=(action, src, dst, w, h, m), daemon=True).start()

    def work(self, action, src, dst, w, h, m):
        def prog(i, n):
            self.after(0, lambda: (self.progress.config(maximum=n, value=i),
                                   self.status.config(text=f"Resizing page {i}/{n}…")))
        try:
            n = resize_pdf(src, dst, w, h, m, self.mode.get(), self.auto_orient.get(), prog)
            self.after(0, lambda: self.finish(action, dst, n, None))
        except Exception as e:
            self.after(0, lambda: self.finish(action, dst, 0, str(e)))

    def finish(self, action, dst, n, err):
        self.busy = False
        for b in (self.btn_save, self.btn_print, self.btn_preview):
            b.config(state="normal")
        if err:
            self.status.config(text="Failed.")
            messagebox.showerror("PDF Fit & Print", err)
            return
        self.status.config(text=f"Done - {n} page(s) resized.")
        try:
            if action == "print":
                send_to_printer(dst)
                self.status.config(text="Sent to printer.")
            elif action == "preview":
                open_file(dst)
            else:
                if messagebox.askyesno("Saved", f"Saved to:\n{dst}\n\nOpen it now?"):
                    open_file(dst)
        except Exception:
            # Direct printing not available - open in viewer so user can press Ctrl+P
            open_file(dst)
            messagebox.showinfo("Print", "Opened the resized PDF - use Print from the viewer.")


if __name__ == "__main__":
    App().mainloop()
