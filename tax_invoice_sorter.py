import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import threading
import multiprocessing
import os
import time
import re
import json
import hashlib
from pypdf import PdfReader, PdfWriter
import pandas as pd

S3M_PATTERN = re.compile(r'S3M\d+')
CACHE_DIR = os.path.expanduser("~/.tax_invoice_sorter_cache")

def get_pdf_hash(pdf_path):
    stat = os.stat(pdf_path)
    key = f"{pdf_path}|{stat.st_size}|{stat.st_mtime}"
    return hashlib.md5(key.encode()).hexdigest()

def load_cache(pdf_hash):
    os.makedirs(CACHE_DIR, exist_ok=True)
    cache_file = os.path.join(CACHE_DIR, f"{pdf_hash}.json")
    if os.path.exists(cache_file):
        try:
            with open(cache_file, "r") as f:
                return json.load(f)
        except:
            return None
    return None

def save_cache(pdf_hash, data):
    os.makedirs(CACHE_DIR, exist_ok=True)
    cache_file = os.path.join(CACHE_DIR, f"{pdf_hash}.json")
    with open(cache_file, "w") as f:
        json.dump(data, f)

def _scan_chunk(args):
    pdf_path, page_indices = args
    reader = PdfReader(pdf_path)
    results = {}
    for i in page_indices:
        text = (reader.pages[i].extract_text() or "").upper()
        all_s3m = S3M_PATTERN.findall(text)
        results[str(i)] = all_s3m
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

def scan_pdf(pdf_path, progress_cb=None):
    """Scan PDF, return page_data. Pakai cache kalau ada."""
    pdf_hash = get_pdf_hash(pdf_path)
    cached = load_cache(pdf_hash)
    if cached:
        if progress_cb:
            progress_cb(70, 100, "Cache ditemukan! Skip scanning...")
        return cached, True

    reader = PdfReader(pdf_path)
    total_pages = len(reader.pages)

    num_workers = min(multiprocessing.cpu_count(), 8)
    chunk_size  = max(1, total_pages // (num_workers * 4))
    chunks = [list(range(i, min(i + chunk_size, total_pages)))
              for i in range(0, total_pages, chunk_size)]

    args_list = [(pdf_path, chunk) for chunk in chunks]
    page_data = {}
    chunks_done = 0

    if progress_cb:
        progress_cb(0, 100, f"Scanning PDF dengan {num_workers} core CPU... (0%)")

    with multiprocessing.Pool(processes=num_workers) as pool:
        for chunk_result in pool.imap_unordered(_scan_chunk, args_list):
            page_data.update(chunk_result)
            chunks_done += 1
            pages_done = min(chunks_done * chunk_size, total_pages)
            pct = int((pages_done / max(total_pages, 1)) * 70)
            if progress_cb:
                progress_cb(pct, 100,
                    f"Scanning PDF... {pages_done}/{total_pages} halaman ({pct}%)")

    save_cache(pdf_hash, page_data)
    return page_data, False

def build_output(pdf_path, page_data, invoice_list, output_path, progress_cb=None):
    invoice_set = set(inv.upper() for inv in invoice_list)

    if progress_cb:
        progress_cb(72, 100, "Mengelompokkan bundle faktur... (72%)")

    all_s3m_pages = sorted([int(k) for k, v in page_data.items() if len(v) > 0])

    page_invoice_map = {}
    for k, s3m_list in page_data.items():
        for s in s3m_list:
            if s in invoice_set:
                page_invoice_map[int(k)] = s
                break

    invoice_pages = {}
    for idx, page_idx in enumerate(all_s3m_pages):
        inv_key = page_invoice_map.get(page_idx)
        if inv_key is None or inv_key not in invoice_set:
            continue
        start = 0 if idx == 0 else all_s3m_pages[idx - 1] + 1
        bundle = list(range(start, page_idx + 1))
        if inv_key not in invoice_pages:
            invoice_pages[inv_key] = bundle

    if progress_cb:
        progress_cb(75, 100, "Menyusun output PDF... (75%)")

    reader = PdfReader(pdf_path)
    writer = PdfWriter()
    found     = 0
    not_found = 0
    missing   = []
    total     = len(invoice_list)

    for idx, inv in enumerate(invoice_list):
        pct = 75 + int((idx / max(total, 1)) * 24)
        if progress_cb:
            progress_cb(pct, 100, f"Menyusun PDF... {idx+1}/{total} ({pct}%)")
        pages = invoice_pages.get(inv.upper())
        if pages:
            for p in pages:
                writer.add_page(reader.pages[p])
            found += 1
        else:
            not_found += 1
            missing.append(inv)

    if progress_cb:
        progress_cb(99, 100, "Menyimpan file output... (99%)")

    with open(output_path, "wb") as f:
        writer.write(f)

    return found, not_found, missing


# =====================================================================
#  MODE NOTA RETUR
#  Nota Retur Coretax: halaman pertama memuat "Nomor: RET + 15 digit",
#  halaman lanjutan (lampiran barang) ada DI BELAKANGNYA, dan setiap
#  halaman punya penanda "X dari Y" di pojok bawah.
# =====================================================================

RET_PATTERN       = re.compile(r'RET\d{15}(?!\d)')
RET_NOMOR_PATTERN = re.compile(r'NOMOR\s*:?\s*(RET[\d\s]{15,25})')
PAGE_MARK_PATTERN = re.compile(r'(?<!\d)(\d{1,3})\s*DARI\s*(\d{1,3})(?!\d)')
RETUR_CACHE_VER   = "retur-v1"


def normalize_ret(value):
    """'ret 0426 0002 2912281' -> 'RET042600022912281' (atau None)."""
    m = RET_PATTERN.search(re.sub(r'\s+', '', str(value or '')).upper())
    return m.group(0) if m else None


def parse_retur_page(text):
    """Ambil No. RET dan penanda 'X dari Y' dari teks satu halaman."""
    up = (text or "").upper()
    ret = None
    # 1) Paling ketat: "Nomor: RET" + tepat 15 digit apa adanya
    m = re.search(r'NOMOR\s*:?\s*(RET\d{15})(?!\d)', up)
    if m:
        ret = m.group(1)
    # 2) RET + 15 digit di mana saja di halaman
    if not ret:
        m = RET_PATTERN.search(up)
        if m:
            ret = m.group(0)
    # 3) Teks terpotong spasi ("RET 0426 ..."): gabung spasi dalam satu baris saja
    if not ret:
        m = RET_NOMOR_PATTERN.search(up)
        if m:
            ret = normalize_ret(m.group(1).split("\n")[0])
    mark = None
    for a, b in PAGE_MARK_PATTERN.findall(up):
        x, y = int(a), int(b)
        if 1 <= x <= y <= 200:
            mark = [x, y]          # ambil yang terakhir (pojok bawah)
    return {"ret": ret, "mark": mark}


def get_retur_cache_key(pdf_path):
    # Pakai nama + ukuran + waktu ubah, BUKAN lokasi folder,
    # jadi PDF yang dipindah folder tetap kena cache.
    stat = os.stat(pdf_path)
    key = f"{RETUR_CACHE_VER}|{os.path.basename(pdf_path)}|{stat.st_size}|{int(stat.st_mtime)}"
    return "retur_" + hashlib.md5(key.encode()).hexdigest()


def open_pdf_lowmem(pdf_path):
    """
    Buka PDF tanpa memuat seluruh file ke RAM.
    (PdfReader(path) membaca SELURUH file ke memori; dengan file handle,
    pypdf membaca langsung dari disk seperlunya. Penting untuk PDF ratusan MB
    yang dibaca banyak core sekaligus.)
    """
    fh = open(pdf_path, "rb")
    return PdfReader(fh), fh


def _scan_chunk_retur(args):
    pdf_path, page_indices = args
    reader, fh = open_pdf_lowmem(pdf_path)
    results = {}
    try:
        for i in page_indices:
            try:
                text = reader.pages[i].extract_text() or ""
            except Exception:
                text = ""
            results[str(i)] = parse_retur_page(text)
    finally:
        fh.close()
    return results


def scan_pdf_retur(pdf_path, progress_cb=None):
    """Scan Nota Retur per halaman pakai semua core. Pakai cache kalau ada."""
    key = get_retur_cache_key(pdf_path)
    cached = load_cache(key)
    if cached and cached.get("ver") == RETUR_CACHE_VER:
        if progress_cb:
            progress_cb(70, 100, "Cache ditemukan! Skip scanning...")
        return cached["pages"], True

    reader, fh = open_pdf_lowmem(pdf_path)
    total_pages = len(reader.pages)
    fh.close()
    del reader

    num_workers = max(1, min(multiprocessing.cpu_count(), 8))
    chunk_size  = max(1, min(200, total_pages // (num_workers * 4) or 1))
    chunks = [list(range(i, min(i + chunk_size, total_pages)))
              for i in range(0, total_pages, chunk_size)]
    args_list = [(pdf_path, c) for c in chunks]

    pages = [None] * total_pages
    pages_done = 0
    if progress_cb:
        progress_cb(0, 100, f"Scanning Nota Retur dengan {num_workers} core CPU... (0%)")

    with multiprocessing.Pool(processes=num_workers) as pool:
        for chunk_result in pool.imap_unordered(_scan_chunk_retur, args_list):
            for k, v in chunk_result.items():
                pages[int(k)] = v
            pages_done += len(chunk_result)
            pct = int((pages_done / max(total_pages, 1)) * 70)
            if progress_cb:
                progress_cb(pct, 100,
                    f"Scanning Nota Retur... {pages_done}/{total_pages} halaman ({pct}%)")

    save_cache(key, {"ver": RETUR_CACHE_VER, "pages": pages})
    return pages, False


def group_retur_pages(pages):
    """
    Kelompokkan halaman jadi nota.
    Halaman ber-No. RET memulai nota baru bila nomornya beda dari nota
    sebelumnya, atau bertanda '1 dari Y'. Halaman tanpa No. RET / nomor sama
    dianggap lampiran nota sebelumnya. Setiap nota diverifikasi pakai
    penanda 'X dari Y'; yang tidak cocok diberi catatan (perlu cek).
    Return: (dict ret -> nota pertama, list halaman yatim, list nota dobel)
    """
    notas, orphans, cur = [], [], None

    def close(n):
        issues = []
        marks = n["marks"]
        known = [m for m in marks if m]
        if known:
            y = known[0][1]
            if any(m[1] != y for m in known):
                issues.append('Penanda "X dari Y" tidak konsisten')
            if len(n["pages"]) != y:
                issues.append(f"Jumlah halaman {len(n['pages'])}, seharusnya {y}")
            for i, m in enumerate(marks):
                if m and m[0] != i + 1:
                    issues.append(f'Halaman ke-{i+1} bertanda "{m[0]} dari {m[1]}"')
            if any(m is None for m in marks):
                issues.append('Ada halaman tanpa penanda "X dari Y"')
        else:
            issues.append('Tidak ada penanda "X dari Y" untuk verifikasi')
        n["issues"] = list(dict.fromkeys(issues))
        notas.append(n)

    for i, p in enumerate(pages):
        p = p or {"ret": None, "mark": None}
        ret, mark = p.get("ret"), p.get("mark")
        starts_new = ret and (cur is None or ret != cur["ret"] or (mark and mark[0] == 1))
        if starts_new:
            if cur:
                close(cur)
            cur = {"ret": ret, "pages": [i], "marks": [mark]}
        elif cur:
            cur["pages"].append(i)
            cur["marks"].append(mark)
        else:
            orphans.append(i)
    if cur:
        close(cur)

    by_ret, dup_in_pdf = {}, []
    for n in notas:
        if n["ret"] in by_ret:
            dup_in_pdf.append(n)
        else:
            by_ret[n["ret"]] = n
    for n in dup_in_pdf:
        first = by_ret[n["ret"]]
        first["issues"].append(
            f"No. RET ini muncul lagi di halaman {n['pages'][0]+1}; yang dipakai yang pertama")
    return by_ret, orphans, dup_in_pdf


def extract_retur_numbers_from_excel(excel_path, column_name):
    """Ambil No. RET unik (urut kemunculan pertama) + jumlah kemunculannya."""
    df = pd.read_excel(excel_path, dtype=str)
    df.columns = df.columns.str.strip()
    if column_name not in df.columns:
        raise ValueError(f"Kolom '{column_name}' tidak ditemukan.")
    order, counts, first_row, invalid = [], {}, {}, []
    for idx, raw in enumerate(df[column_name].tolist()):
        if raw is None or (isinstance(raw, float) and pd.isna(raw)) or not str(raw).strip():
            continue
        ret = normalize_ret(raw)
        if not ret:
            invalid.append((idx + 2, str(raw).strip()))
            continue
        if ret not in counts:
            order.append(ret)
            counts[ret] = 0
            first_row[ret] = idx + 2      # +2: baris header + index 0
        counts[ret] += 1
    return order, counts, first_row, invalid


def guess_retur_column(excel_path, cols):
    """Tebak kolom yang berisi No. RET."""
    try:
        df = pd.read_excel(excel_path, dtype=str, nrows=50)
        df.columns = df.columns.str.strip()
        best, best_hits = None, 0
        for c in cols:
            hits = sum(1 for v in df[c].tolist() if normalize_ret(v))
            if hits > best_hits:
                best, best_hits = c, hits
        if best:
            return best
    except Exception:
        pass
    for c in cols:
        if "RET" in c.upper():
            return c
    return cols[0] if cols else ""


def build_output_retur(pdf_path, pages, excel_info, output_path, report_path,
                       progress_cb=None):
    order, counts, first_row, invalid = excel_info
    if progress_cb:
        progress_cb(72, 100, "Memasangkan halaman lampiran tiap nota... (72%)")
    by_ret, orphans, dup_in_pdf = group_retur_pages(pages)

    if progress_cb:
        progress_cb(75, 100, "Menyusun output PDF... (75%)")
    reader, fh = open_pdf_lowmem(pdf_path)
    writer = PdfWriter()
    found = not_found = need_check = out_pages = 0
    rows, missing, check_list = [], [], []
    total = len(order)

    for idx, ret in enumerate(order):
        pct = 75 + int((idx / max(total, 1)) * 22)
        if progress_cb and (idx % 10 == 0 or idx == total - 1):
            progress_cb(pct, 100, f"Menyusun PDF... {idx+1}/{total} ({pct}%)")
        n = by_ret.get(ret)
        notes = []
        if counts[ret] > 1:
            notes.append(f"Dobel {counts[ret]}x di Excel, diambil sekali")
        if n:
            for p in n["pages"]:
                writer.add_page(reader.pages[p])
            found += 1
            out_pages += len(n["pages"])
            status = "Perlu cek" if n["issues"] else "Ditemukan"
            if n["issues"]:
                need_check += 1
                check_list.append((ret, "; ".join(n["issues"])))
                notes = n["issues"] + notes
            rows.append([idx + 1, ret, status, f"{n['pages'][0]+1}-{n['pages'][-1]+1}"
                         if len(n["pages"]) > 1 else str(n["pages"][0] + 1),
                         len(n["pages"]), out_pages - len(n["pages"]) + 1,
                         first_row[ret], "; ".join(notes)])
        else:
            not_found += 1
            missing.append(ret)
            rows.append([idx + 1, ret, "Tidak ditemukan", "", 0, "", first_row[ret],
                         "; ".join(notes)])

    for r, v in invalid:
        rows.append(["", v, "Format tidak valid", "", 0, "", r,
                     "Bukan format RET + 15 digit, dilewati"])

    if progress_cb:
        progress_cb(98, 100, "Menyimpan file output... (98%)")
    try:
        if found:
            with open(output_path, "wb") as f:
                writer.write(f)
    finally:
        fh.close()

    report = pd.DataFrame(rows, columns=[
        "Urut", "No RET", "Status", "Halaman di PDF sumber", "Jumlah halaman",
        "Mulai di halaman output", "Baris di Excel", "Catatan"])
    extra = []
    if orphans:
        extra.append(["Halaman tanpa No. RET di awal PDF",
                      ", ".join(str(i + 1) for i in orphans[:50]) + (" ..." if len(orphans) > 50 else "")])
    for n in dup_in_pdf:
        extra.append([f"{n['ret']} dobel di PDF", f"halaman {n['pages'][0]+1} (tidak dipakai)"])
    with pd.ExcelWriter(report_path, engine="openpyxl") as xw:
        report.to_excel(xw, sheet_name="Laporan", index=False)
        if extra:
            pd.DataFrame(extra, columns=["Info", "Detail"]).to_excel(
                xw, sheet_name="Catatan PDF", index=False)
        ws = xw.sheets["Laporan"]
        for col, width in zip("ABCDEFGH", (7, 22, 16, 20, 15, 22, 13, 70)):
            ws.column_dimensions[col].width = width

    return {"missing": missing, "check_list": check_list,
            "found": found, "not_found": not_found, "need_check": need_check,
            "total": total, "pages": out_pages, "invalid": len(invalid)}


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
        self.mode_var   = tk.StringVar(value="faktur")
        self._build()
        self._apply_mode()
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

        self._card_mode(pad)
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
        outer._title = tk.Label(hdr, text=f"  {title}", bg=CARDBG, fg=TEXT,
                                font=("Helvetica", 12, "bold"))
        outer._title.pack(side="left")
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

    # ---------- Pilihan jenis dokumen ----------
    MODE_TEXT = {
        "faktur": {
            "pdf_title": "File PDF Faktur Pajak",
            "col_title": "Pilih kolom No. Invoice",
            "col_hint":  "App otomatis baca semua nilai dari kolom ini dan cocokkan ke PDF.",
            "info": "Output hanya berisi faktur yang ada di Excel, diurutkan sesuai Excel. "
                    "PDF yang sudah pernah diproses akan di-cache otomatis sehingga "
                    "proses berikutnya jauh lebih cepat.",
            "button": "Mulai Sortir Faktur Pajak",
        },
        "retur": {
            "pdf_title": "File PDF Nota Retur",
            "col_title": "Pilih kolom No. Retur",
            "col_hint":  "Nilai berformat RET + 15 digit dibaca dari kolom ini. "
                         "No. RET yang dobel diambil sekali (posisi pertama).",
            "info": "Output berisi nota retur sesuai urutan Excel, lengkap dengan halaman "
                    "lampirannya (dicek pakai penanda \"X dari Y\"). Laporan .xlsx berisi "
                    "No. RET yang tidak ditemukan / perlu cek ikut disimpan di folder output.",
            "button": "Mulai Sortir Nota Retur",
        },
    }

    def _card_mode(self, parent):
        card = self._make_card(parent, "0", "Jenis Dokumen")
        row = tk.Frame(card, bg=CARDBG)
        row.pack(fill="x", padx=16, pady=(0, 14))
        row.columnconfigure((0, 1), weight=1, uniform="m")
        self._mode_btns = {}
        for col, (val, label) in enumerate((("faktur", "Faktur Pajak"),
                                            ("retur",  "Nota Retur"))):
            b = tk.Radiobutton(row, text=label, value=val, variable=self.mode_var,
                               indicatoron=0, font=("Helvetica", 11, "bold"),
                               relief="flat", bd=0, cursor="hand2", pady=9,
                               highlightthickness=1, highlightbackground=BORDER,
                               command=self._apply_mode)
            b.grid(row=0, column=col, sticky="ew", padx=(0 if col == 0 else 8, 0))
            self._mode_btns[val] = b

    def _mode(self):
        return self.mode_var.get()

    def _apply_mode(self):
        m = self._mode()
        txt = self.MODE_TEXT[m]
        for val, b in self._mode_btns.items():
            on = (val == m)
            b.configure(bg=ACCENT if on else ZONEBG, fg="white" if on else TEXT2,
                        selectcolor=ACCENT if on else ZONEBG,
                        activebackground="#1D4ED8" if on else BORDER,
                        activeforeground="white" if on else TEXT)
        self.pdf_card._title.configure(text=f"  {txt['pdf_title']}")
        self.col_title.configure(text=txt["col_title"])
        self.col_hint.configure(text=txt["col_hint"])
        self.info_lbl.configure(text=txt["info"])
        if str(self.run_btn["state"]) != "disabled":
            self.run_btn.configure(text=txt["button"])
        self._refresh_pdf_status()
        xl = self.excel_path.get()
        if m == "retur" and xl and self.col_menu["values"]:
            self.col_var.set(guess_retur_column(xl, list(self.col_menu["values"])))

    def _refresh_pdf_status(self):
        path = self.pdf_path.get()
        if not path or not os.path.exists(path):
            return
        size = os.path.getsize(path) / (1024*1024)
        if self._mode() == "retur":
            c = load_cache(get_retur_cache_key(path))
            cached = bool(c and c.get("ver") == RETUR_CACHE_VER)
        else:
            cached = bool(load_cache(get_pdf_hash(path)))
        self._zone_done(self.pdf_zone, "PDF", os.path.basename(path),
                        f"{size:.1f} MB - Cache tersedia (scan instan)" if cached
                        else f"{size:.1f} MB - Belum ada cache (akan scan pertama kali)")

    def _card_pdf(self, parent):
        card = self._make_card(parent, "1", "File PDF Faktur Pajak")
        self.pdf_card = card
        self.pdf_zone = self._make_upload_zone(
            card, "PDF", "Klik untuk pilih file PDF", "Format: .pdf", self._pick_pdf)

    def _pick_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF Files","*.pdf")])
        if not path: return
        self.pdf_path.set(path)
        self._refresh_pdf_status()

    def _card_excel(self, parent):
        card = self._make_card(parent, "2", "Data Excel")
        self.excel_zone = self._make_upload_zone(
            card, "XLS", "Klik untuk pilih file Excel", "Format: .xlsx / .xls",
            self._pick_excel)
        self.col_frame = tk.Frame(card, bg=CARDBG)
        tk.Frame(self.col_frame, bg=BORDER, height=1).pack(fill="x", pady=(4, 10))
        self.col_title = tk.Label(self.col_frame, text="Pilih kolom No. Invoice",
                 bg=CARDBG, fg=TEXT2,
                 font=("Helvetica", 11, "bold"))
        self.col_title.pack(anchor="w", padx=16)
        self.col_hint = tk.Label(self.col_frame,
                 text="App otomatis baca semua nilai dari kolom ini dan cocokkan ke PDF.",
                 bg=CARDBG, fg=MUTED, font=("Helvetica", 10),
                 wraplength=580, justify="left")
        self.col_hint.pack(anchor="w", padx=16, pady=(2,8))
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
        self._zone_done(self.excel_zone, "XLS", os.path.basename(path),
                        f"{rows} baris - {len(cols)} kolom terdeteksi")
        self.col_menu["values"] = cols
        if self._mode() == "retur":
            self.col_var.set(guess_retur_column(path, cols))
        else:
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
        tk.Button(row, text="Pilih Folder",
                  font=("Helvetica", 11), fg=TEXT2, bg="#F3F4F6",
                  activebackground=BORDER, relief="flat", cursor="hand2",
                  padx=12, pady=6, command=self._pick_output).pack(side="left")
        info = tk.Frame(card, bg=INFOBG, highlightthickness=1,
                        highlightbackground="#BFDBFE")
        info.pack(fill="x", padx=16, pady=(0,16))
        self.info_lbl = tk.Label(info,
                 text="Output hanya berisi faktur yang ada di Excel, diurutkan "
                      "sesuai Excel. PDF yang sudah pernah diproses akan di-cache "
                      "otomatis sehingga proses berikutnya jauh lebih cepat.",
                 bg=INFOBG, fg=INFOFG, font=("Helvetica", 10),
                 wraplength=560, justify="left")
        self.info_lbl.pack(padx=12, pady=8)

    def _pick_output(self):
        path = filedialog.askdirectory()
        if path:
            self.output_dir.set(path)
            self._path_entry.configure(fg=TEXT)

    def _btn_run(self, parent):
        self.run_btn = tk.Button(
            parent,
            text="Mulai Sortir Faktur Pajak",
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
        self.stat_found    = self._stat(stats, "-", "Cocok & diurutkan", SUCCESS, 0)
        self.stat_notfound = self._stat(stats, "-", "Tidak ditemukan",   WARN,    1)
        self.stat_total    = self._stat(stats, "-", "Total di Excel",    MUTED,   2)

        # Daftar nomor yang tidak ditemukan / perlu cek
        self.list_frame = tk.Frame(inner, bg=CARDBG)
        lhdr = tk.Frame(self.list_frame, bg=CARDBG)
        lhdr.pack(fill="x", pady=(14, 6))
        self.list_title = tk.Label(lhdr, text="Tidak ditemukan di PDF", bg=CARDBG,
                                   fg=TEXT2, font=("Helvetica", 11, "bold"))
        self.list_title.pack(side="left")
        self.copy_btn = tk.Button(lhdr, text="Salin daftar", font=("Helvetica", 10),
                                  fg=TEXT2, bg="#F3F4F6", activebackground=BORDER,
                                  relief="flat", cursor="hand2", padx=10, pady=3,
                                  command=self._copy_list)
        self.copy_btn.pack(side="right")
        box = tk.Frame(self.list_frame, bg=CARDBG, highlightthickness=1,
                       highlightbackground=BORDER)
        box.pack(fill="x")
        self.list_text = tk.Text(box, height=10, wrap="word", relief="flat",
                                 font=("Consolas", 10), bg=ZONEBG, fg=TEXT,
                                 padx=10, pady=8, borderwidth=0)
        lsb = tk.Scrollbar(box, orient="vertical", command=self.list_text.yview)
        self.list_text.configure(yscrollcommand=lsb.set)
        lsb.pack(side="right", fill="y")
        self.list_text.pack(side="left", fill="both", expand=True)
        self.list_text.tag_configure("bad",  foreground="#B91C1C")
        self.list_text.tag_configure("warn", foreground="#B45309")
        self.list_text.tag_configure("ok",   foreground=DONEFG)
        self.list_text.tag_configure("head", foreground=TEXT2, font=("Consolas", 10, "bold"))
        # Text di dalam canvas: scroll mouse di kotak ini menggulung daftar, bukan halaman
        self.list_text.bind("<Enter>", lambda e: self._canvas.unbind_all("<MouseWheel>"))
        self.list_text.bind("<Leave>", lambda e: self._canvas.bind_all("<MouseWheel>",
            lambda ev: self._canvas.yview_scroll(int(-1*(ev.delta/120)), "units")))
        self.list_text.bind("<MouseWheel>",
            lambda e: (self.list_text.yview_scroll(int(-1*(e.delta/120)), "units"), "break")[1])
        self._list_plain = ""
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

    def _show_list(self, missing, check_list=None, label="No."):
        """Tampilkan nomor yang tidak ditemukan (dan perlu cek) di bawah statistik."""
        check_list = check_list or []
        t = self.list_text
        t.configure(state="normal")
        t.delete("1.0", "end")
        lines = []
        if not missing and not check_list:
            t.insert("end", "Semua nomor di Excel ditemukan di PDF.\n", "ok")
            lines.append("Semua nomor di Excel ditemukan di PDF.")
        if missing:
            h = f"TIDAK DITEMUKAN ({len(missing)})"
            t.insert("end", h + "\n", "head"); lines.append(h)
            for v in missing:
                s = f"{v} tidak ditemukan"
                t.insert("end", s + "\n", "bad"); lines.append(s)
        if check_list:
            if missing:
                t.insert("end", "\n"); lines.append("")
            h = f"PERLU CEK MANUAL ({len(check_list)})"
            t.insert("end", h + "\n", "head"); lines.append(h)
            for v, why in check_list:
                s = f"{v} perlu cek: {why}"
                t.insert("end", s + "\n", "warn"); lines.append(s)
        t.configure(state="disabled")
        self._list_plain = "\n".join(lines)
        n = len(missing)
        self.list_title.configure(
            text=f"Tidak ditemukan di PDF: {n}" + (f"  |  Perlu cek: {len(check_list)}" if check_list else ""))
        self.list_frame.pack(fill="x")
        self.after(100, lambda: self._canvas.yview_moveto(1.0))

    def _missing_preview(self, missing, limit=10):
        """Potongan daftar untuk jendela 'Selesai' (daftar lengkap ada di aplikasi)."""
        if not missing:
            return ""
        s = "\nTidak ditemukan:\n" + "".join(f"  {v}\n" for v in missing[:limit])
        if len(missing) > limit:
            s += f"  ... dan {len(missing) - limit} lainnya (lihat daftar di aplikasi)\n"
        return s + "\n"

    def _hide_list(self):
        self.list_frame.pack_forget()
        self._list_plain = ""

    def _copy_list(self):
        if not self._list_plain:
            return
        self.clipboard_clear()
        self.clipboard_append(self._list_plain)
        self.copy_btn.configure(text="Tersalin")
        self.after(1500, lambda: self.copy_btn.configure(text="Salin daftar"))

    def _animate_dots(self):
        if not self._loading:
            self.prog_loading.config(text="")
            return
        self._dot_count = (self._dot_count + 1) % 4
        dots = "*" * self._dot_count + "." * (3 - self._dot_count)
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
            messagebox.showwarning("Perhatian",
                "Pilih kolom No. Retur." if self._mode() == "retur"
                else "Pilih kolom No. Invoice."); return
        if not outd:
            messagebox.showwarning("Perhatian", "Pilih folder output terlebih dahulu."); return
        self.run_btn.config(state="disabled", text="Sedang memproses...",
                            bg=GRAY, fg="#E5E7EB")
        self.prog_bar["value"] = 0
        self.stat_found.config(text="-")
        self.stat_notfound.config(text="-")
        self.stat_total.config(text="-")
        self._hide_list()
        self._loading   = True
        self._dot_count = 0
        self._animate_dots()
        for b in self._mode_btns.values():
            b.configure(state="disabled")
        target = self._worker_retur if self._mode() == "retur" else self._worker
        threading.Thread(target=target,
                         args=(pdf, xl, col, outd), daemon=True).start()

    def _worker(self, pdf_path, excel_path, col, output_dir):
        try:
            self._status("Membaca data Excel...")
            invoice_list = extract_invoice_numbers_from_excel(excel_path, col)
            total = len(invoice_list)
            self.after(0, lambda: self.stat_total.config(text=str(total)))

            ts   = time.strftime("%Y%m%d_%H%M%S")
            outp = os.path.join(output_dir, f"Faktur_Sorted_{ts}.pdf")

            def cb(pct, _, msg):
                self.after(0, lambda p=pct, m=msg: (
                    self.prog_bar.configure(value=p),
                    self.prog_status.configure(text=m)
                ))

            page_data, from_cache = scan_pdf(pdf_path, cb)
            if from_cache:
                cb(71, 100, "Cache loaded! Langsung matching...")

            found, not_found, missing = build_output(
                pdf_path, page_data, invoice_list, outp, cb)

            self.after(0, lambda: self._done(found, not_found, total, outp, from_cache, missing))
        except Exception as e:
            self.after(0, lambda err=str(e): self._error(err))

    def _worker_retur(self, pdf_path, excel_path, col, output_dir):
        try:
            self._status("Membaca data Excel...")
            excel_info = extract_retur_numbers_from_excel(excel_path, col)
            total = len(excel_info[0])
            if total == 0:
                raise ValueError(f"Tidak ada No. RET (format RET + 15 digit) di kolom '{col}'.")
            self.after(0, lambda: self.stat_total.config(text=str(total)))

            ts   = time.strftime("%Y%m%d_%H%M%S")
            outp = os.path.join(output_dir, f"Retur_Sorted_{ts}.pdf")
            repp = os.path.join(output_dir, f"Retur_Sorted_{ts}_laporan.xlsx")

            def cb(pct, _, msg):
                self.after(0, lambda p=pct, m=msg: (
                    self.prog_bar.configure(value=p),
                    self.prog_status.configure(text=m)
                ))

            pages, from_cache = scan_pdf_retur(pdf_path, cb)
            if from_cache:
                cb(71, 100, "Cache loaded! Langsung matching...")

            res = build_output_retur(pdf_path, pages, excel_info, outp, repp, cb)
            self.after(0, lambda: self._done_retur(res, outp, repp, from_cache))
        except Exception as e:
            self.after(0, lambda err=str(e): self._error(err))

    def _reset_run_btn(self):
        self.run_btn.config(state="normal",
                            text=self.MODE_TEXT[self._mode()]["button"],
                            bg=ACCENT, fg="white")
        for b in self._mode_btns.values():
            b.configure(state="normal")

    def _done_retur(self, res, output_path, report_path, from_cache):
        self._loading = False
        self.prog_bar["value"] = 100
        cache_note = "Dari cache" if from_cache else "Cache disimpan untuk next run"
        extra = f" - {res['need_check']} perlu cek" if res["need_check"] else ""
        self.prog_status.config(text=f"Selesai! {cache_note}{extra}")
        self.prog_loading.config(text="")
        self.stat_found.config(text=str(res["found"]))
        self.stat_notfound.config(text=str(res["not_found"]))
        self.stat_total.config(text=str(res["total"]))
        self._reset_run_btn()
        self._show_list(res["missing"], res["check_list"])
        lines = [f"- Cocok & diurutkan : {res['found']} nota ({res['pages']} halaman)",
                 f"- Tidak ditemukan   : {res['not_found']}",
                 f"- Total di Excel    : {res['total']} No. RET unik"]
        if res["need_check"]:
            lines.append(f"- PERLU CEK         : {res['need_check']} nota (lihat laporan)")
        if res["invalid"]:
            lines.append(f"- Format tidak valid: {res['invalid']} baris dilewati")
        out_line = (f"Output:\n{output_path}\n" if res["found"]
                    else "PDF tidak dibuat karena tidak ada yang cocok.\n")
        if messagebox.askyesno("Selesai!",
            "Nota retur berhasil diurutkan!\n\n" + "\n".join(lines) + "\n" +
            self._missing_preview(res["missing"]) +
            f"- {cache_note}\n\n{out_line}Laporan:\n{report_path}\n\n"
            "Buka folder output sekarang?"):
            os.startfile(os.path.dirname(output_path))

    def _status(self, msg):
        self.after(0, lambda: self.prog_status.config(text=msg))

    def _done(self, found, not_found, total, output_path, from_cache, missing=None):
        self._loading = False
        self.prog_bar["value"] = 100
        cache_note = "Dari cache" if from_cache else "Cache disimpan untuk next run"
        self.prog_status.config(text=f"Selesai! {cache_note}")
        self.prog_loading.config(text="")
        self.stat_found.config(text=str(found))
        self.stat_notfound.config(text=str(not_found))
        self.stat_total.config(text=str(total))
        self._reset_run_btn()
        self._show_list(missing or [])
        if messagebox.askyesno("Selesai!",
            f"Faktur berhasil diurutkan!\n\n"
            f"- Cocok & diurutkan : {found}\n"
            f"- Tidak ditemukan   : {not_found}\n"
            f"- Total di Excel    : {total}\n"
            + self._missing_preview(missing or []) +
            f"- {cache_note}\n\n"
            f"Output:\n{output_path}\n\n"
            f"Buka folder output sekarang?"):
            os.startfile(os.path.dirname(output_path))

    def _error(self, msg):
        self._loading = False
        self.prog_loading.config(text="")
        self.prog_status.config(text="Terjadi error.")
        self._reset_run_btn()
        messagebox.showerror("Error", f"Terjadi kesalahan:\n\n{msg}")


if __name__ == "__main__":
    multiprocessing.freeze_support()
    app = App()
    app.mainloop()
