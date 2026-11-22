"""
PDF Fit & Print - resize every page of a PDF to a custom paper size, then save or print.

Printing on Windows is done directly to the chosen printer (no Adobe dialog):
each page is rendered and sent at the custom paper size you entered.

Requires: pip install pypdf pypdfium2 pillow pywin32   (pywin32 is Windows only)
Run:      python pdf_fit_print.py
Build:    pyinstaller --onefile --windowed --collect-all pypdfium2 --collect-all pypdfium2_raw pdf_fit_print.py
"""
import os
import sys
import queue
import tempfile
import threading
import traceback
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from pypdf import PdfReader, PdfWriter, Transformation
from pypdf.generic import RectangleObject

IS_WIN = sys.platform.startswith("win")
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

LOG_FILE = os.path.join(tempfile.gettempdir(), "pdf_fit_print_log.txt")


# ------------------------------------------------------------------ resizing
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
        # Attach the page to the writer FIRST (required by pypdf 7+), then edit it
        page = writer.add_page(page)
        page.transfer_rotation_to_content()  # bake /Rotate into the content
        box = page.cropbox  # visible area (defaults to mediabox)
        w, h = float(box.width), float(box.height)
        left, bottom = float(box.left), float(box.bottom)

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

        op = (Transformation().translate(-left, -bottom)
              .scale(sx, sy).translate(tx, ty))
        page.add_transformation(op)

        page.mediabox = RectangleObject([0, 0, W, H])
        page.cropbox = RectangleObject([0, 0, W, H])
        page.trimbox = RectangleObject([0, 0, W, H])
        page.bleedbox = RectangleObject([0, 0, W, H])
        page.artbox = RectangleObject([0, 0, W, H])

        if progress:
            progress(idx, total)

    with open(dst, "wb") as f:
        writer.write(f)
    return total


def booklet_pdf(src, dst, sheet_w_pt, sheet_h_pt, margin_pt=0.0, gutter_pt=0.0,
                mode="fit", progress=None, which="all", reverse_backs=False,
                rotate_backs=False):
    """Impose `src` as a folded booklet: two pages per sheet side, in folding order.

    Each sheet side is sheet_w x sheet_h; the left and right halves each hold one page.
    Pages are pushed toward the fold so any spare space ends up on the outer edges.
    Page count is padded with blanks to a multiple of 4.

    which = "all"   -> every sheet side in order (front1, back1, front2, back2, ...)
            "front" -> only the fronts (sheet sides 1, 3, 5, ...)
            "back"  -> only the backs  (sheet sides 2, 4, 6, ...)
    reverse_backs -> output the backs in reverse order
    rotate_backs  -> turn the back sides 180 degrees (for how you re-feed the paper)
    """
    reader = PdfReader(src)
    if reader.is_encrypted:
        reader.decrypt("")
    pages = list(reader.pages)
    n = len(pages)
    N = ((n + 3) // 4) * 4
    pages += [None] * (N - n)

    # saddle-stitch order: front (N-1-2i, 2i), back (2i+1, N-2-2i)
    sides = []
    for i in range(N // 4):
        sides.append((N - 1 - 2 * i, 2 * i))
        sides.append((2 * i + 1, N - 2 - 2 * i))

    W, H = sheet_w_pt, sheet_h_pt
    half = W / 2.0
    slots = (  # (x_min, x_max) for left and right page
        (margin_pt, half - gutter_pt / 2.0),
        (half + gutter_pt / 2.0, W - margin_pt),
    )
    y_min, y_max = margin_pt, H - margin_pt

    jobs = [(i, side) for i, side in enumerate(sides)
            if which == "all" or (which == "front" and i % 2 == 0)
            or (which == "back" and i % 2 == 1)]
    if reverse_backs and which == "back":
        jobs.reverse()

    writer = PdfWriter()
    for k, (side_idx, side) in enumerate(jobs, 1):
        sheet = writer.add_blank_page(width=W, height=H)
        for slot_idx, page_no in enumerate(side):
            page = pages[page_no]
            if page is None:
                continue
            box = page.cropbox
            w, h = float(box.width), float(box.height)
            left, bottom = float(box.left), float(box.bottom)
            rot = (page.rotation or 0) % 360

            # visible size after /Rotate, and matrix that applies it
            if rot == 90:
                vw, vh, ctm = h, w, (0, -1, 1, 0, 0, w)
            elif rot == 180:
                vw, vh, ctm = w, h, (-1, 0, 0, -1, w, h)
            elif rot == 270:
                vw, vh, ctm = h, w, (0, 1, -1, 0, h, 0)
            else:
                vw, vh, ctm = w, h, (1, 0, 0, 1, 0, 0)

            x0, x1 = slots[slot_idx]
            aw, ah = max(x1 - x0, 1), max(y_max - y_min, 1)
            if mode == "stretch":
                sx, sy = aw / vw, ah / vh
                tx, ty = x0, y_min
            else:
                sc = min(aw / vw, ah / vh)
                sx = sy = sc
                ty = y_min + (ah - vh * sc) / 2
                # left page hugs the fold (right edge of its slot); right page hugs it too
                tx = (x1 - vw * sc) if slot_idx == 0 else x0

            op = (Transformation().translate(-left, -bottom)
                  .transform(Transformation(ctm=ctm))
                  .scale(sx, sy).translate(tx, ty))
            sheet.merge_transformed_page(page, op)
        if rotate_backs and side_idx % 2 == 1:
            sheet.rotate(180)
        if progress:
            progress(k, len(jobs))

    with open(dst, "wb") as f:
        writer.write(f)
    return len(jobs), n, N


# ------------------------------------------------------------------ printing
def list_printers():
    """Return (printer names, default printer) - Windows only."""
    if not IS_WIN:
        return [], ""
    try:
        import win32print
        names = [p[2] for p in win32print.EnumPrinters(
            win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
        return names, win32print.GetDefaultPrinter()
    except Exception:
        return [], ""


def print_pdf_windows(path, printer, use_custom_size, progress=None, duplex=False):
    """Render each page of `path` and print it straight to `printer`.

    If use_custom_size is True the page's own size is requested from the printer
    driver as a custom paper size (works only if the driver supports it).
    """
    import win32con
    import win32gui
    import win32print
    import win32ui
    import pypdfium2 as pdfium
    from PIL import ImageWin

    DMPAPER_USER = 256
    pdf = pdfium.PdfDocument(path)
    total = len(pdf)

    hprinter = win32print.OpenPrinter(printer)
    try:
        dm = win32print.GetPrinter(hprinter, 2)["pDevMode"]

        def set_page_devmode(w_pt, h_pt):
            if use_custom_size:
                w_mm, h_mm = w_pt * 25.4 / 72, h_pt * 25.4 / 72
                dm.PaperSize = DMPAPER_USER
                dm.PaperWidth = int(round(min(w_mm, h_mm) * 10))   # tenths of a mm
                dm.PaperLength = int(round(max(w_mm, h_mm) * 10))
                dm.Fields |= (win32con.DM_PAPERSIZE | win32con.DM_PAPERWIDTH
                              | win32con.DM_PAPERLENGTH)
            dm.Orientation = 2 if w_pt > h_pt else 1  # landscape / portrait
            dm.Fields |= win32con.DM_ORIENTATION
            if duplex:
                dm.Duplex = 3  # DMDUP_HORIZONTAL = flip on the short edge
                dm.Fields |= win32con.DM_DUPLEX

        w0, h0 = pdf[0].get_size()
        set_page_devmode(w0, h0)
        handle = win32gui.CreateDC("WINSPOOL", printer, dm)
        hdc = win32ui.CreateDCFromHandle(handle)
        hdc.StartDoc("PDF Fit & Print")
        try:
            for i in range(total):
                page = pdf[i]
                w_pt, h_pt = page.get_size()
                if i > 0:
                    set_page_devmode(w_pt, h_pt)
                    try:
                        win32gui.ResetDC(hdc.GetSafeHdc(), dm)
                    except Exception:
                        pass  # keep previous page settings
                hdc.StartPage()

                horz = hdc.GetDeviceCaps(win32con.HORZRES)
                vert = hdc.GetDeviceCaps(win32con.VERTRES)
                dpi = min(hdc.GetDeviceCaps(win32con.LOGPIXELSX) or 300, 300)
                scale = dpi / 72.0
                # keep memory sane on very large sheets (max ~60 megapixels)
                px = (w_pt * scale) * (h_pt * scale)
                if px > 60e6:
                    scale *= (60e6 / px) ** 0.5
                img = page.render(scale=scale).to_pil().convert("RGB")

                r = min(horz / img.width, vert / img.height)
                dw, dh = int(img.width * r), int(img.height * r)
                x0, y0 = (horz - dw) // 2, (vert - dh) // 2
                ImageWin.Dib(img).draw(hdc.GetHandleOutput(), (x0, y0, x0 + dw, y0 + dh))
                hdc.EndPage()
                if progress:
                    progress(i + 1, total)
            hdc.EndDoc()
        except BaseException:
            try:
                hdc.AbortDoc()
            except Exception:
                pass
            raise
        finally:
            hdc.DeleteDC()
    finally:
        win32print.ClosePrinter(hprinter)
    return total


def print_fallback(path):
    """No pywin32 / not Windows: hand the file to the system."""
    if IS_WIN:
        os.startfile(path, "print")  # default PDF app, default printer
    else:
        subprocess.check_call(["lp", path])


def open_file(path):
    if IS_WIN:
        os.startfile(path)
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


# ------------------------------------------------------------------ GUI
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF Fit & Print")
        self.geometry("560x760")
        self.resizable(False, False)

        self.src = tk.StringVar()
        self.preset = tk.StringVar(value="Custom")
        self.width = tk.StringVar(value="600")
        self.height = tk.StringVar(value="215")
        self.unit = tk.StringVar(value="mm")
        self.margin = tk.StringVar(value="0")
        self.gutter = tk.StringVar(value="0")
        self.booklet = tk.BooleanVar(value=True)
        self.duplex = tk.BooleanVar(value=False)
        self.which = tk.StringVar(value="all")
        self.rev_backs = tk.BooleanVar(value=False)
        self.rot_backs = tk.BooleanVar(value=False)
        self.mode = tk.StringVar(value="fit")
        self.auto_orient = tk.BooleanVar(value=True)
        self.printer = tk.StringVar()
        self.custom_paper = tk.BooleanVar(value=True)
        self.busy = False
        self.q = queue.Queue()
        self._build()
        self.after(100, self.poll)  # main thread checks messages from the worker

    def _build(self):
        p = {"padx": 10, "pady": 3}
        f = ttk.Frame(self)
        f.pack(fill="both", expand=True, padx=8, pady=8)

        ttk.Label(f, text="PDF file:").grid(row=0, column=0, sticky="w", **p)
        ttk.Entry(f, textvariable=self.src, width=40).grid(row=0, column=1, columnspan=2, sticky="ew", **p)
        ttk.Button(f, text="Browse…", command=self.pick).grid(row=0, column=3, **p)

        box = ttk.LabelFrame(f, text="Sheet size (for a booklet: both pages together)")
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
        ttk.Label(box, text="Middle gap:").grid(row=3, column=0, sticky="w", **p)
        ttk.Entry(box, textvariable=self.gutter, width=9).grid(row=3, column=1, **p)
        ttk.Label(box, text="(0 = pages touch at the fold)").grid(row=3, column=2, columnspan=2,
                                                                  sticky="w", **p)

        opt = ttk.LabelFrame(f, text="How to fit")
        opt.grid(row=2, column=0, columnspan=4, sticky="ew", **p)
        ttk.Checkbutton(opt, text="Convert to booklet (2 pages per sheet, folding order)",
                        variable=self.booklet).pack(anchor="w", padx=10, pady=2)
        ttk.Radiobutton(opt, text="Fit inside (keep proportions, centred)",
                        variable=self.mode, value="fit").pack(anchor="w", padx=10, pady=2)
        ttk.Radiobutton(opt, text="Stretch to fill (may distort)",
                        variable=self.mode, value="stretch").pack(anchor="w", padx=10, pady=2)
        ttk.Checkbutton(opt, text="Auto-rotate sheet to match each page's orientation",
                        variable=self.auto_orient).pack(anchor="w", padx=10, pady=2)

        man = ttk.LabelFrame(f, text="Booklet sheets to print (for manual front/back)")
        man.grid(row=3, column=0, columnspan=4, sticky="ew", **p)
        ttk.Radiobutton(man, text="All", variable=self.which,
                        value="all").grid(row=0, column=0, sticky="w", padx=10)
        ttk.Radiobutton(man, text="Fronts only (odd: 1, 3, 5…)", variable=self.which,
                        value="front").grid(row=0, column=1, sticky="w", padx=10)
        ttk.Radiobutton(man, text="Backs only (even: 2, 4, 6…)", variable=self.which,
                        value="back").grid(row=0, column=2, sticky="w", padx=10)
        ttk.Checkbutton(man, text="Reverse order of backs",
                        variable=self.rev_backs).grid(row=1, column=0, columnspan=2,
                                                      sticky="w", padx=10, pady=2)
        ttk.Checkbutton(man, text="Turn backs 180°",
                        variable=self.rot_backs).grid(row=1, column=2, sticky="w", padx=10, pady=2)

        prn = ttk.LabelFrame(f, text="Printer")
        prn.grid(row=4, column=0, columnspan=4, sticky="ew", **p)
        names, default = list_printers()
        self.printer.set(default or (names[0] if names else ""))
        ttk.Label(prn, text="Printer:").grid(row=0, column=0, sticky="w", **p)
        ttk.Combobox(prn, textvariable=self.printer, values=names,
                     state="readonly" if names else "disabled",
                     width=42).grid(row=0, column=1, sticky="w", **p)
        ttk.Checkbutton(prn, text="Ask the printer for this exact paper size",
                        variable=self.custom_paper).grid(row=1, column=0, columnspan=2,
                                                         sticky="w", padx=10, pady=2)
        ttk.Checkbutton(prn, text="Print double-sided (flip on short edge)",
                        variable=self.duplex).grid(row=2, column=0, columnspan=2,
                                                   sticky="w", padx=10, pady=2)
        if not names:
            ttk.Label(prn, text="(No printer list - will use the system default)"
                      ).grid(row=3, column=0, columnspan=2, sticky="w", padx=10)

        self.progress = ttk.Progressbar(f, mode="determinate")
        self.progress.grid(row=5, column=0, columnspan=4, sticky="ew", **p)
        self.status = ttk.Label(f, text="Choose a PDF to begin.")
        self.status.grid(row=6, column=0, columnspan=4, sticky="w", **p)

        bf = ttk.Frame(f)
        bf.grid(row=7, column=0, columnspan=4, pady=8)
        self.btn_save = ttk.Button(bf, text="Save resized PDF…", command=lambda: self.run("save"))
        self.btn_print = ttk.Button(bf, text="Print", command=lambda: self.run("print"))
        self.btn_preview = ttk.Button(bf, text="Preview", command=lambda: self.run("preview"))
        for b in (self.btn_save, self.btn_preview, self.btn_print):
            b.pack(side="left", padx=6)
        f.columnconfigure(1, weight=1)

    def set_buttons(self, enabled):
        state = "normal" if enabled else "disabled"
        for b in (self.btn_save, self.btn_print, self.btn_preview):
            b.config(state=state)

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
            g = float(self.gutter.get() or 0) * PT_PER[self.unit.get()]
            if g < 0:
                raise ValueError("Middle gap cannot be negative.")
        except ValueError as e:
            messagebox.showwarning("PDF Fit & Print", str(e))
            return

        if action == "save":
            base = os.path.splitext(os.path.basename(src))[0]
            tag = {"front": "_fronts", "back": "_backs"}.get(self.which.get(), "_resized") \
                if self.booklet.get() else "_resized"
            dst = filedialog.asksaveasfilename(defaultextension=".pdf",
                                               initialfile=f"{base}{tag}.pdf",
                                               filetypes=[("PDF files", "*.pdf")])
            if not dst:
                return
        else:
            dst = os.path.join(tempfile.gettempdir(), "pdf_fit_print_temp.pdf")

        self.busy = True
        self.set_buttons(False)
        self.progress.config(value=0)
        self.status.config(text="Working… large PDFs can take a minute.")
        opts = dict(mode=self.mode.get(), auto=self.auto_orient.get(),
                    printer=self.printer.get(), custom=self.custom_paper.get(),
                    booklet=self.booklet.get(), gutter=g, duplex=self.duplex.get(),
                    which=self.which.get(), rev=self.rev_backs.get(),
                    rot=self.rot_backs.get())
        threading.Thread(target=self.work, args=(action, src, dst, w, h, m, opts),
                         daemon=True).start()

    # ---- worker thread: never touches Tk, only puts messages on the queue ----
    def work(self, action, src, dst, w, h, m, o):
        try:
            cb = lambda i, t: self.q.put(("progress", f"Building sheet {i}/{t}…", i, t))
            if o["booklet"]:
                n, orig, padded = booklet_pdf(src, dst, w, h, m, o["gutter"], o["mode"], cb,
                                              o["which"], o["rev"], o["rot"])
            else:
                n = resize_pdf(src, dst, w, h, m, o["mode"], o["auto"], cb)
            if action == "print":
                self.q.put(("progress", "Sending to printer…", 0, n))
                used_fallback = False
                if IS_WIN and o["printer"]:
                    try:
                        print_pdf_windows(
                            dst, o["printer"], o["custom"],
                            lambda i, t: self.q.put(("progress", f"Printing page {i}/{t}…", i, t)),
                            duplex=o["duplex"] and o["booklet"] and o["which"] == "all")
                    except ImportError:
                        print_fallback(dst)
                        used_fallback = True
                else:
                    print_fallback(dst)
                    used_fallback = True
                self.q.put(("done", action, dst, n, used_fallback))
            else:
                self.q.put(("done", action, dst, n, False))
        except BaseException as e:
            try:
                with open(LOG_FILE, "w", encoding="utf-8") as f:
                    f.write(traceback.format_exc())
            except Exception:
                pass
            self.q.put(("error", f"{type(e).__name__}: {e}\n\nDetails saved to:\n{LOG_FILE}"))

    # ---- main thread: reads messages and updates the window ----
    def poll(self):
        try:
            while True:
                msg = self.q.get_nowait()
                if msg[0] == "progress":
                    _, text, i, t = msg
                    self.progress.config(maximum=max(t, 1), value=i)
                    self.status.config(text=text)
                elif msg[0] == "done":
                    self.finish(*msg[1:])
                elif msg[0] == "error":
                    self.busy = False
                    self.set_buttons(True)
                    self.status.config(text="Failed.")
                    messagebox.showerror("PDF Fit & Print", msg[1])
        except queue.Empty:
            pass
        self.after(100, self.poll)

    def finish(self, action, dst, n, used_fallback):
        self.busy = False
        self.set_buttons(True)
        try:
            if action == "print":
                self.status.config(text=f"Sent {n} page(s) to the printer.")
                if used_fallback:
                    messagebox.showinfo("Print", "Sent to your default PDF app for printing.")
            elif action == "preview":
                self.status.config(text=f"Done - {n} page(s) resized.")
                open_file(dst)
            else:
                self.status.config(text=f"Done - {n} page(s) resized.")
                if messagebox.askyesno("Saved", f"Saved to:\n{dst}\n\nOpen it now?"):
                    open_file(dst)
        except Exception as e:
            messagebox.showerror("PDF Fit & Print", str(e))


if __name__ == "__main__":
    App().mainloop()
