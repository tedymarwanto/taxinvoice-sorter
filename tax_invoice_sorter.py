import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import threading
import multiprocessing
import os
import time
import re
from pypdf import PdfReader, PdfWriter
import pandas as pd

def _scan_chunk(args):
    pdf_path, page_indices, invoice_set = args
    reader = PdfReader(pdf_path)
    results = {}
    pattern = re.compile(r'S3M\d+')
    for i in page_indices:
        text = (reader.pages[i].extract_text() or "").upper()
        found_excel = None
        for inv in invoice_set:
            if inv in text:
                found_excel = inv
                break
        all_s3m = pattern.findall(text)
        results[i] = (found_excel, len(all_s3m) > 0)
    return results

def get_excel_columns(excel_path):
    df = pd.read_excel(excel_path, nrows=0, dtype=str)
    return [c.strip() for c in df.columns.tolist()]

def extract_invoice_numbers_from_excel(excel_path, column_name):
    df = pd.read_excel(excel_path, dtype=str)
    df.columns = df.columns.str.strip()
    if column_name not in df.columns:
        raise ValueError(f"Kolom '{column_name}' tidak ditemukan.")
    raw = df[column_name].dropna().str.strip().tolist()
    seen = set()
    unique = []
    for inv in raw:
        if inv.upper() not in seen:
            seen.add(inv.upper())
            unique.append(inv)
    return unique

def scan_and_build(pdf_path, invoice_list, output_path, progress_cb=None):
    invoice_set = set(inv.upper() for inv in invoice_list)

    reader = PdfReader(pdf_path)
    total_pages = len(reader.pages)

    num_workers = min(multiprocessing.cpu_count(), 8)
    chunk_size  = max(1, total_pages // num_workers)
    chunks = [list(range(i, min(i + chunk_size, total_pages)))
              for i in range(0, total_pages, chunk_size)]

    args_list = [(pdf_path, chunk, invoice_set) for chunk in chunks]

    # page_data[i] = (invoice_key_or_None, has_any_s3m)
    page_data = {}
    completed = 0

    if progress_cb:
        progress_cb(0, total_pages, f"Menyiapkan {num_workers} core CPU...")

    with multiprocessing.Pool(processes=num_workers) as pool:
        for chunk_result in pool.imap_unordered(_scan_chunk, args_list):
            page_data.update(chunk_result)
            completed += chunk_size
            if progress_cb:
                progress_cb(
                    min(completed, total_pages),
                    total_pages,
                    f"Scanning... {min(completed, total_pages)}/{total_pages} halaman"
                )

    if progress_cb:
        progress_cb(total_pages, total_pages, "Mengelompokkan faktur...")

    # Kumpulkan semua halaman yang punya S3M apapun (sebagai batas bundle)
    all_s3m_pages = sorted([i for i in range(total_pages)
                            if page_data.get(i, (None, False))[1]])

    # Build invoice_pages:
    # Untuk tiap halaman yang S3M nya ada di Excel,
    # ambil bundle dari setelah S3M-apapun sebelumnya sampai halaman ini
    invoice_pages = {}

    for idx, page_idx in enumerate(all_s3m_pages):
        inv_key = page_data.get(page_idx, (None, False))[0]
        if inv_key is None or inv_key not in invoice_set:
            continue

        # Batas awal = setelah S3M apapun sebelumnya
        start = 0 if idx == 0 else all_s3m_pages[idx - 1] + 1
        bundle = list(range(start, page_idx + 1))

        if inv_key not in invoice_pages:
            invoice_pages[inv_key] = bundle

    # Susun output sesuai urutan Excel
    writer    = PdfWriter()
    found     = 0
    not_found = 0
    total     = len(invoice_list)

    for idx, inv in enumerate(invoice_list):
        if progress_cb:
            pct = 50 + int((idx / max(total, 1)) * 50)
            progress_cb(pct, 100, f"Menyusun PDF... {idx+1}/{total}")
        pages = invoice_pages.get(inv.upper())
        if pages:
            for p in pages:
                writer.add_page(reader.pages[p])
            found += 1
        else:
            not_found += 1

    with open(output_path, "wb") as f:
        writer.write(f)

    return found, not_found


ACCENT  = "#2563EB"
SUCCESS = "#16A34A"
WARN    = "#D97706"
MUTED   = "#6B7280"
BORDER  = "#D1D5DB"
TEXT    = "#111827"
TEXT2   = "#374151"
BG      = "#F3F4F6"
CARDBG  = "#FFFFFF"
INFOBG  = "#EFF6FF"
INFOFG  = "#1E40AF"
DONEBG  = "#F0FDF4"
DONEFG  = "#15803D"
ZONEBG  = "#F9FAFB"
GRAY    = "#9CA3AF"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Tax Invoice Sorter")
        self.configure(bg=BG)
        self.resizable(False, False)
        self.pdf_path   = tk.StringVar()
        self.excel_path = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.col_var    = tk.StringVar()
        self._build()
        self._center(720, 780)

    def _center(self, w, h):
        self.update_idletasks()
        x = (self.winfo_screenwidth()  - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _build(self):
        bar = tk.Frame(self, bg="#E5E7EB", height=38)
        bar.pack(fill="x")
        bar.pack_propagate(False)
        tk.Label(bar, text="Tax Invoice Sorter", bg="#E5E7EB", fg="#555",
                 font=("Helvetica", 12)).pack(expand=True)

        self._canvas = tk.Canvas(self, bg=BG, highlightthickness=0)
        sb = tk.Scrollbar(self, orient="vertical", command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self._canvas.pack(side="left", fill="both", expand=True)

        self._sf = tk.Frame(self._canvas, bg=BG)
        self._win = self._canvas.create_window((0,0), window=self._sf, anchor="nw")
        self._sf.bind("<Configure>",
            lambda e: self._canvas.configure(scrollregion=self._canvas.bbox("all")))
        self._canvas.bind("<Configure>",
            lambda e: self._canvas.itemconfig(self._win, width=e.width))
        self._canvas.bind_all("<MouseWheel>",
            lambda e: self._canvas.yview_scroll(int(-1*(e.delta/120)), "units"))

        pad = tk.Frame(self._sf, bg=BG)
        pad.pack(fill="both", expand=True, padx=24, pady=20)

        self._card_pdf(pad)
        self._card_excel(pad)
        self._card_output(pad)
        self._btn_run(pad)
        self._section_progress(pad)

    def _make_card(self, parent, step, title):
        outer = tk.Frame(parent, bg=CARDBG, highlightthickness=1,
                         highlightbackground=BORDER)
        outer.pack(fill="x", pady=(0, 12))
        hdr = tk.Frame(outer, bg=CARDBG)
        hdr.pack(fill="x", padx=16, pady=(14, 10))
        badge = tk.Canvas(hdr, width=24, height=24, bg=CARDBG, highlightthickness=0)
        badge.pack(side="left")
        badge.create_oval(0, 0, 24, 24, fill=ACCENT, outline=ACCENT)
        badge.create_text(12, 12, text=step, fill="white",
                          font=("Helvetica", 10, "bold"))
        tk.Label(hdr, text=f"  {title}", bg=CARDBG, fg=TEXT,
                 font=("Helvetica", 12, "bold")).pack(side="left")
        return outer

    def _make_upload_zone(self, parent, icon, label, sublabel, command):
        zone = tk.Frame(parent, bg=ZONEBG, highlightthickness=1,
                        highlightbackground=BORDER, cursor="hand2")
        zone.pack(fill="x", padx=16, pady=(0, 14))
        inner = tk.Frame(zone, bg=ZONEBG)
        inner.pack(fill="x", padx=16, pady=12)
        icon_lbl = tk.Label(inner, text=icon, bg=ZONEBG, font=("Helvetica", 22))
        icon_lbl.pack(side="left", padx=(0, 12))
        tf = tk.Frame(inner, bg=ZONEBG)
        tf.pack(side="left", fill="x", expand=True)
        main_lbl = tk.Label(tf, text=label, bg=ZONEBG, fg=ACCENT,
                            font=("Helvetica", 11, "bold"))
        main_lbl.pack(anchor="w")
        sub_lbl = tk.Label(tf, text=sublabel, bg=ZONEBG, fg=MUTED,
                           font=("Helvetica", 10))
        sub_lbl.pack(anchor="w")
        for w in (zone, inner, icon_lbl, tf, main_lbl, sub_lbl):
            w.bind("<Button-1>", lambda e, cmd=command: cmd())
        zone._icon  = icon_lbl
        zone._main  = main_lbl
        zone._sub   = sub_lbl
        zone._inner = inner
        return zone

    def _zone_done(self, zone, icon, fname, sub):
        for w in (zone, zone._inner, zone._icon, zone._main, zone._sub):
            w.configure(bg=DONEBG)
        zone.configure(highlightbackground="#86EFAC")
        zone._icon.configure(text=icon)
        zone._main.configure(text=fname, fg=DONEFG)
        zone._sub.configure(text=sub, fg=DONEFG)

    def _card_pdf(self, parent):
        card = self._make_card(parent, "1", "File PDF Faktur Pajak")
        self.pdf_zone = self._make_upload_zone(
            card, "📄", "Klik untuk pilih file PDF", "Format: .pdf", self._pick_pdf)

    def _pick_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF Files","*.pdf")])
        if not path: return
        self.pdf_path.set(path)
        size = os.path.getsize(path) / (1024*1024)
        self._zone_done(self.pdf_zone, "📄", os.path.basename(path),
                        f"{size:.1f} MB · berhasil dibaca ✓")

    def _card_excel(self, parent):
        card = self._make_card(parent, "2", "Data Excel")
        self.excel_zone = self._make_upload_zone(
            card, "📊", "Klik untuk pilih file Excel", "Format: .xlsx / .xls",
            self._pick_excel)
        self.col_frame = tk.Frame(card, bg=CARDBG)
        tk.Frame(self.col_frame, bg=BORDER, height=1).pack(fill="x", pady=(4, 10))
        tk.Label(self.col_frame, text="Pilih kolom No. Invoice",
                 bg=CARDBG, fg=TEXT2,
                 font=("Helvetica", 11, "bold")).pack(anchor="w", padx=16)
        tk.Label(self.col_frame,
                 text="App otomatis baca semua nilai dari kolom ini dan cocokkan ke PDF.",
                 bg=CARDBG, fg=MUTED, font=("Helvetica", 10),
                 wraplength=580, justify="left").pack(anchor="w", padx=16, pady=(2,8))
        self.col_menu = ttk.Combobox(self.col_frame, textvariable=self.col_var,
                                     state="readonly", font=("Helvetica", 11))
        self.col_menu.pack(fill="x", padx=16, pady=(0,14))

    def _pick_excel(self):
        path = filedialog.askopenfilename(
            filetypes=[("Excel Files","*.xlsx *.xls")])
        if not path: return
        try:
            cols = get_excel_columns(path)
        except Exception as e:
            messagebox.showerror("Error", f"Gagal membaca Excel:\n{e}"); return
        self.excel_path.set(path)
        df = pd.read_excel(path, dtype=str)
        rows = len(df)
        self._zone_done(self.excel_zone, "📊", os.path.basename(path),
                        f"{rows} baris · {len(cols)} kolom terdeteksi ✓")
        self.col_menu["values"] = cols
        self.col_var.set(cols[0] if cols else "")
        self.col_frame.pack(fill="x")

    def _card_output(self, parent):
        card = self._make_card(parent, "3", "Lokasi Output")
        row = tk.Frame(card, bg=CARDBG)
        row.pack(fill="x", padx=16, pady=(0,10))
        self._path_entry = tk.Entry(row, textvariable=self.output_dir,
                                    font=("Helvetica", 11), fg=MUTED,
                                    bg=ZONEBG, relief="flat",
                                    highlightthickness=1,
                                    highlightbackground=BORDER,
                                    highlightcolor=ACCENT)
        self._path_entry.pack(side="left", fill="x", expand=True, ipady=7, padx=(0,8))
        tk.Button(row, text="📁  Pilih Folder",
                  font=("Helvetica", 11), fg=TEXT2, bg="#F3F4F6",
                  activebackground=BORDER, relief="flat", cursor="hand2",
                  padx=12, pady=6, command=self._pick_output).pack(side="left")
        info = tk.Frame(card, bg=INFOBG, highlightthickness=1,
                        highlightbackground="#BFDBFE")
        info.pack(fill="x", padx=16, pady=(0,16))
        tk.Label(info,
                 text="ℹ️   Output hanya berisi faktur yang ada di Excel, "
                      "diurutkan sesuai urutan Excel. Faktur multi-halaman "
                      "otomatis dipasangkan sampai ketemu barcode S3M.",
                 bg=INFOBG, fg=INFOFG, font=("Helvetica", 10),
                 wraplength=560, justify="left").pack(padx=12, pady=8)

    def _pick_output(self):
        path = filedialog.askdirectory()
        if path:
            self.output_dir.set(path)
            self._path_entry.configure(fg=TEXT)

    def _btn_run(self, parent):
        self.run_btn = tk.Button(
            parent,
            text="▲  Mulai Sortir Faktur Pajak",
            font=("Helvetica", 13, "bold"),
            fg="white", bg=ACCENT,
            activeforeground="white",
            activebackground="#1D4ED8",
            relief="flat", cursor="hand2",
            pady=14,
            command=self._start
        )
        self.run_btn.pack(fill="x", pady=(4, 0))

    def _section_progress(self, parent):
        self.prog_card = tk.Frame(parent, bg=CARDBG, highlightthickness=1,
                                  highlightbackground=BORDER)
        self.prog_card.pack(fill="x", pady=(12, 0))
        inner = tk.Frame(self.prog_card, bg=CARDBG)
        inner.pack(fill="x", padx=20, pady=16)
        self.prog_status = tk.Label(inner, text="Menunggu proses dimulai...",
                                    bg=CARDBG, fg=MUTED, font=("Helvetica", 11))
        self.prog_status.pack(anchor="w")
        self.prog_loading = tk.Label(inner, text="", bg=CARDBG, fg=ACCENT,
                                     font=("Helvetica", 11, "bold"))
        self.prog_loading.pack(anchor="w")
        style = ttk.Style()
        style.configure("Blue.Horizontal.TProgressbar",
                        troughcolor=BORDER, background=ACCENT,
                        thickness=10, borderwidth=0)
        self.prog_bar = ttk.Progressbar(inner, style="Blue.Horizontal.TProgressbar",
                                         mode="determinate", maximum=100)
        self.prog_bar.pack(fill="x", pady=(8, 14))
        stats = tk.Frame(inner, bg=CARDBG)
        stats.pack(fill="x")
        stats.columnconfigure((0,1,2), weight=1, uniform="s")
        self.stat_found    = self._stat(stats, "–", "Cocok & diurutkan", SUCCESS, 0)
        self.stat_notfound = self._stat(stats, "–", "Tidak ditemukan",   WARN,    1)
        self.stat_total    = self._stat(stats, "–", "Total di Excel",    MUTED,   2)
        self._loading   = False
        self._dot_count = 0

    def _stat(self, parent, num, label, color, col):
        f = tk.Frame(parent, bg=ZONEBG, highlightthickness=1,
                     highlightbackground=BORDER)
        f.grid(row=0, column=col, sticky="ew", padx=(0 if col==0 else 6, 0))
        n = tk.Label(f, text=num, bg=ZONEBG, fg=color,
                     font=("Helvetica", 22, "bold"))
        n.pack(padx=12, pady=(10,2))
        tk.Label(f, text=label, bg=ZONEBG, fg=MUTED,
                 font=("Helvetica", 10)).pack(padx=12, pady=(0,10))
        return n

    def _animate_dots(self):
        if not self._loading:
            self.prog_loading.config(text="")
            return
        self._dot_count = (self._dot_count + 1) % 4
        dots = "●" * self._dot_count + "○" * (3 - self._dot_count)
        self.prog_loading.config(text=dots)
        self.after(400, self._animate_dots)

    def _start(self):
        pdf  = self.pdf_path.get()
        xl   = self.excel_path.get()
        col  = self.col_var.get()
        outd = self.output_dir.get()
        if not pdf:
            messagebox.showwarning("Perhatian", "Pilih file PDF terlebih dahulu."); return
        if not xl:
            messagebox.showwarning("Perhatian", "Pilih file Excel terlebih dahulu."); return
        if not col:
            messagebox.showwarning("Perhatian", "Pilih kolom No. Invoice."); return
        if not outd:
            messagebox.showwarning("Perhatian", "Pilih folder output terlebih dahulu."); return
        self.run_btn.config(state="disabled", text="⏳  Sedang memproses...",
                            bg=GRAY, fg="#E5E7EB")
        self.prog_bar["value"] = 0
        self.stat_found.config(text="–")
        self.stat_notfound.config(text="–")
        self.stat_total.config(text="–")
        self._loading   = True
        self._dot_count = 0
        self._animate_dots()
        threading.Thread(target=self._worker,
                         args=(pdf, xl, col, outd), daemon=True).start()

    def _worker(self, pdf_path, excel_path, col, output_dir):
        try:
            self._status("Membaca data Excel...")
            invoice_list = extract_invoice_numbers_from_excel(excel_path, col)
            total = len(invoice_list)
            self.after(0, lambda: self.stat_total.config(text=str(total)))
            ts   = time.strftime("%Y%m%d_%H%M%S")
            outp = os.path.join(output_dir, f"Faktur_Sorted_{ts}.pdf")

            def cb(cur, tot, msg):
                pct = int((cur / max(tot, 1)) * 100)
                self.after(0, lambda p=min(pct,100), m=msg: (
                    self.prog_bar.configure(value=p),
                    self.prog_status.configure(text=m)
                ))

            found, not_found = scan_and_build(pdf_path, invoice_list, outp, cb)
            self.after(0, lambda: self._done(found, not_found, total, outp))
        except Exception as e:
            self.after(0, lambda: self._error(str(e)))

    def _status(self, msg):
        self.after(0, lambda: self.prog_status.config(text=msg))

    def _done(self, found, not_found, total, output_path):
        self._loading = False
        self.prog_bar["value"] = 100
        self.prog_status.config(text="✅  Selesai! File PDF berhasil dibuat.")
        self.prog_loading.config(text="")
        self.stat_found.config(text=str(found))
        self.stat_notfound.config(text=str(not_found))
        self.stat_total.config(text=str(total))
        self.run_btn.config(state="normal",
                            text="▲  Mulai Sortir Faktur Pajak",
                            bg=ACCENT, fg="white")
        if messagebox.askyesno("Selesai!",
            f"✅ Faktur berhasil diurutkan!\n\n"
            f"• Cocok & diurutkan : {found}\n"
            f"• Tidak ditemukan   : {not_found}\n"
            f"• Total di Excel    : {total}\n\n"
            f"Output:\n{output_path}\n\n"
            f"Buka folder output sekarang?"):
            os.system(f'open "{os.path.dirname(output_path)}"')

    def _error(self, msg):
        self._loading = False
        self.prog_loading.config(text="")
        self.prog_status.config(text="❌  Terjadi error.")
        self.run_btn.config(state="normal",
                            text="▲  Mulai Sortir Faktur Pajak",
                            bg=ACCENT, fg="white")
        messagebox.showerror("Error", f"Terjadi kesalahan:\n\n{msg}")


if __name__ == "__main__":
    multiprocessing.freeze_support()
    app = App()
    app.mainloop()
