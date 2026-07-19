"""
AutoAMLE
=========
🚀 AutoAMLE – Automating AMLE Profiles, LNREL Relations, and actAmle Activation
Developed by Abdelrahman Mohamed Galal

Filtering logic per CSV:
  AMLEPR  → source-band cells only (per band-mapping row)
             AND only the amlePrId whose targetCarrierFreq in MDB matches chosen target
  LNREL   → ALL cells in site where mrbtsId == ecgiAdjEnbId  (no band filter)
  LC      → ALL cells in site  (no band filter at all)
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import os
import threading
from datetime import datetime

try:
    import pyodbc
    _PYODBC_OK = True
except ImportError:
    _PYODBC_OK = False

# ─────────────────────────────────────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────────────────────────────────────

AUTHOR = "Developed by Abdelrahman Mohamed Galal"

BAND_FREQ = {
    "L18":  1675,
    "L9":   3622,
    "L21":  224,
    "TDD1": 40392,
    "TDD2": 40590,
}

BAND_CELLS = {
    "L18":  lambda c: (1  <= c <= 9)  or (99  <= c <= 199),
    "L9":   lambda c: (10 <= c <= 19),
    "L21":  lambda c: (20 <= c <= 29) or (200 <= c <= 299),
    "TDD1": lambda c: (30 <= c <= 39),
    "TDD2": lambda c: (40 <= c <= 49),
}

BANDS = list(BAND_FREQ.keys())

# ─────────────────────────────────────────────────────────────────────────────
# MDB helpers
# ─────────────────────────────────────────────────────────────────────────────

def _conn_str(mdb_path):
    return (r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};"
            f"Dbq={mdb_path};")

def list_mdb_tables(mdb_path):
    conn = pyodbc.connect(_conn_str(mdb_path))
    tables = [r.table_name for r in conn.cursor().tables(tableType="TABLE")]
    conn.close()
    return tables

def read_mdb_table(mdb_path, table_name):
    conn = pyodbc.connect(_conn_str(mdb_path))
    df = pd.read_sql(f"SELECT * FROM [{table_name}]", conn)
    conn.close()
    df.columns = [c.strip() for c in df.columns]
    return df

def find_table(tables, *keywords):
    for kw in keywords:
        for t in tables:
            if kw.lower() in t.lower():
                return t
    return None

def get_col(df, name):
    col = next((c for c in df.columns if c.lower() == name.lower()), None)
    if col is None:
        raise KeyError(f"Column '{name}' not found. Available: {list(df.columns)}")
    return col

# ─────────────────────────────────────────────────────────────────────────────
# Core processing
# ─────────────────────────────────────────────────────────────────────────────

def process(mdb_path, mrbts_ids, band_rows, output_dir, log):

    log("=" * 68)
    log(f"  AutoAMLE")
    log(f"  {datetime.now().strftime('%Y-%m-%d  %H:%M:%S')}")
    log(f"  {AUTHOR}")
    log("=" * 68)
    log(f"  MDB    : {mdb_path}")
    log(f"  Sites  : {', '.join(mrbts_ids)}")
    log(f"  Rows   : {len(band_rows)}")
    for i, r in enumerate(band_rows, 1):
        log(f"    [{i:02d}] {r['source']:5s} → {r['target']:5s}  "
            f"cacHeadroom={r['cacHeadroom']}  "
            f"deltaCac={r['deltaCac']}  "
            f"maxCacThreshold={r['maxCacThreshold']}")
    log("=" * 68)

    if not _PYODBC_OK:
        log("  ERROR: pyodbc not installed."); return False

    # ── Load tables ───────────────────────────────────────────────────────────
    log("\n▶ Scanning MDB tables …")
    try:
        tables = list_mdb_tables(mdb_path)
        log(f"  Found {len(tables)} table(s): {tables}")
    except Exception as e:
        log(f"  ERROR: {e}"); return False

    amlepr_tbl = find_table(tables, "AMLEPR", "AMLE")
    lnrel_tbl  = find_table(tables, "LNREL")
    lc_tbl     = find_table(tables, "LTE_LNCEL_LC", "LNCEL_LC", "_LC")

    missing = [n for n, v in [("AMLEPR", amlepr_tbl),
                               ("LNREL",  lnrel_tbl),
                               ("LC",     lc_tbl)] if v is None]
    if missing:
        log(f"  ERROR: Cannot find tables: {missing}")
        return False

    log(f"  AMLEPR → '{amlepr_tbl}'")
    log(f"  LNREL  → '{lnrel_tbl}'")
    log(f"  LC     → '{lc_tbl}'")

    try:
        df_amlepr_full = read_mdb_table(mdb_path, amlepr_tbl)
        df_lnrel_full  = read_mdb_table(mdb_path, lnrel_tbl)
        df_lc_full     = read_mdb_table(mdb_path, lc_tbl)
    except Exception as e:
        log(f"  ERROR reading tables: {e}"); return False

    log(f"  Loaded — AMLEPR:{len(df_amlepr_full)}  "
        f"LNREL:{len(df_lnrel_full)}  LC:{len(df_lc_full)}")

    # ── Site filter ───────────────────────────────────────────────────────────
    ids = [int(x) for x in mrbts_ids]

    def site_filter(df, label):
        col = next((c for c in df.columns if c.lower() == "mrbtsid"), None)
        if not col:
            raise KeyError(f"mrbtsId column not found in {label}")
        return df[df[col].astype(int).isin(ids)].copy()

    try:
        df_amlepr_s = site_filter(df_amlepr_full, "AMLEPR")
        df_lnrel_s  = site_filter(df_lnrel_full,  "LNREL")
        df_lc_s     = site_filter(df_lc_full,     "LC")
    except KeyError as e:
        log(f"  ERROR: {e}"); return False

    log(f"  After site filter — AMLEPR:{len(df_amlepr_s)}  "
        f"LNREL:{len(df_lnrel_s)}  LC:{len(df_lc_s)}")

    amlepr_rows = []
    lnrel_rows  = []
    lc_rows     = []

    # ── AMLEPR: source-band cells + freq match ────────────────────────────────
    log("\n▶ Building AMLEPR …")
    for brow in band_rows:
        src      = brow["source"]
        tgt      = brow["target"]
        headroom = brow["cacHeadroom"]
        delta    = brow["deltaCac"]
        maxcac   = brow["maxCacThreshold"]
        tgt_freq = BAND_FREQ[tgt]
        src_filt = BAND_CELLS[src]

        log(f"  ── {src} → {tgt}  "
            f"cacHeadroom={headroom}  deltaCac={delta}  maxCacThreshold={maxcac}")
        try:
            cel_col  = get_col(df_amlepr_s, "lnCelId")
            freq_col = get_col(df_amlepr_s, "targetCarrierFreq")

            # Step 1 — source band filter
            slice_a = df_amlepr_s[
                df_amlepr_s[cel_col].astype(int).apply(src_filt)].copy()

            # Step 2 — match existing targetCarrierFreq in MDB to chosen target
            slice_a = slice_a[
                slice_a[freq_col].astype(int) == tgt_freq].copy()

            log(f"     band match → freq match ({tgt_freq}): {len(slice_a)} row(s)")
            for _, r2 in slice_a.iterrows():
                amlepr_rows.append({
                    "mrbtsId":           int(r2[get_col(slice_a, "mrbtsId")]),
                    "lnBtsId":           int(r2[get_col(slice_a, "lnBtsId")]),
                    "lnCelId":           int(r2[get_col(slice_a, "lnCelId")]),
                    "amlePrId":          r2[get_col(slice_a, "amlePrId")],
                    "cacHeadroom":       headroom,
                    "deltaCac":          delta,
                    "maxCacThreshold":   maxcac,
                    "targetCarrierFreq": tgt_freq,
                })
        except KeyError as e:
            log(f"     AMLEPR column error: {e}")

    # ── LNREL: all intra-site cells, no band filter ───────────────────────────
    log("\n▶ Building LNREL …")
    try:
        mrb_col  = get_col(df_lnrel_s, "mrbtsId")
        ecgi_col = get_col(df_lnrel_s, "ecgiAdjEnbId")
        lnrel_all = df_lnrel_s[
            df_lnrel_s[mrb_col].astype(int) == df_lnrel_s[ecgi_col].astype(int)
        ].copy()
        log(f"  LNREL intra-site rows: {len(lnrel_all)}")
        for _, r2 in lnrel_all.iterrows():
            lnrel_rows.append({
                "mrbtsId":       int(r2[get_col(lnrel_all, "mrbtsId")]),
                "lnBtsId":       int(r2[get_col(lnrel_all, "lnBtsId")]),
                "lnCelId":       int(r2[get_col(lnrel_all, "lnCelId")]),
                "lnRelId":       r2[get_col(lnrel_all, "lnRelId")],
                "amleAllowed":   1,
                "removeAllowed": 0,
                "ecgiAdjEnbId":  int(r2[ecgi_col]),
                "ecgiLcrId":     r2[get_col(lnrel_all, "ecgiLcrId")],
            })
    except KeyError as e:
        log(f"  LNREL column error: {e}")

    # ── LC: all cells in site, no band filter ─────────────────────────────────
    log("\n▶ Building LC …")
    try:
        log(f"  LC rows for site(s): {len(df_lc_s)}")
        for _, r2 in df_lc_s.iterrows():
            lc_rows.append({
                "mrbtsId": int(r2[get_col(df_lc_s, "mrbtsId")]),
                "lnBtsId": int(r2[get_col(df_lc_s, "lnBtsId")]),
                "lnCelId": int(r2[get_col(df_lc_s, "lnCelId")]),
                "actAmle": 1,
            })
    except KeyError as e:
        log(f"  LC column error: {e}")

    # ── Deduplicate ───────────────────────────────────────────────────────────
    def dedup(rows, keys):
        seen = set(); out = []
        for r in rows:
            k = tuple(r[k] for k in keys)
            if k not in seen:
                seen.add(k); out.append(r)
        return out

    amlepr_rows = dedup(amlepr_rows, ["mrbtsId","lnCelId","amlePrId","targetCarrierFreq"])
    lnrel_rows  = dedup(lnrel_rows,  ["mrbtsId","lnCelId","lnRelId"])
    lc_rows     = dedup(lc_rows,     ["mrbtsId","lnCelId"])

    if not amlepr_rows and not lnrel_rows and not lc_rows:
        log("\n  WARNING: No rows produced. Check site IDs and band selections.")
        return False

    # ── Save CSVs ─────────────────────────────────────────────────────────────
    log("\n▶ Saving CSV files …")

    def save(rows, fname, cols):
        path = os.path.join(output_dir, fname + ".csv")
        pd.DataFrame(rows, columns=cols).to_csv(path, index=False)
        log(f"  ✔  {fname}.csv  →  {len(rows)} row(s)  →  {path}")

    save(amlepr_rows, "A_LTE_MRBTS_LNBTS_LNCEL_AMLEPR",
         ["mrbtsId","lnBtsId","lnCelId","amlePrId",
          "cacHeadroom","deltaCac","maxCacThreshold","targetCarrierFreq"])

    save(lnrel_rows, "A_LTE_MRBTS_LNBTS_LNCEL_LNREL",
         ["mrbtsId","lnBtsId","lnCelId","lnRelId",
          "amleAllowed","removeAllowed","ecgiAdjEnbId","ecgiLcrId"])

    save(lc_rows, "LTE_LNCEL_LC",
         ["mrbtsId","lnBtsId","lnCelId","actAmle"])

    log("\n" + "=" * 68)
    log("  ✔  ALL DONE — 3 CSV files generated.")
    log(f"  {AUTHOR}")
    log("=" * 68)
    return True


# ─────────────────────────────────────────────────────────────────────────────
# GUI colours / fonts
# ─────────────────────────────────────────────────────────────────────────────

BG     = "#0d1117"
PANEL  = "#161b22"
PANEL2 = "#1c2128"
BORDER = "#30363d"
ACCENT = "#00e5ff"
ACCT2  = "#00b4cc"
FG     = "#e6edf3"
FG2    = "#8b949e"
GREEN  = "#3fb950"
RED    = "#f85149"

MONO   = ("Consolas", 10)
MONO_B = ("Consolas", 10, "bold")
MONO_S = ("Consolas", 9)

# ─────────────────────────────────────────────────────────────────────────────
# BandRow widget
# ─────────────────────────────────────────────────────────────────────────────

class BandRow(tk.Frame):

    def __init__(self, parent, row_num, on_delete):
        super().__init__(parent, bg=PANEL2,
                         highlightthickness=1, highlightbackground=BORDER)
        self._num   = row_num
        self._src   = tk.StringVar(value="L18")
        self._tgt   = tk.StringVar(value="TDD1")
        self._head  = tk.StringVar(value="5")
        self._delta = tk.StringVar(value="0")
        self._maxc  = tk.StringVar(value="100")
        self._build(on_delete)

    def _cb(self, var, width=7):
        cb = ttk.Combobox(self, textvariable=var, values=BANDS,
                          width=width, state="readonly", font=MONO)
        cb.bind("<<ComboboxSelected>>", self._update_freq)
        return cb

    def _ent(self, var, width=8):
        return tk.Entry(self, textvariable=var, width=width,
                        bg="#090d11", fg=ACCENT, insertbackground=ACCENT,
                        relief="flat", bd=0, font=MONO_B,
                        highlightthickness=1,
                        highlightbackground=BORDER,
                        highlightcolor=ACCENT)

    def _lbl(self, text, fg=FG2, width=None):
        kw = dict(text=text, font=MONO_S, fg=fg, bg=PANEL2)
        if width:
            kw["width"] = width
        return tk.Label(self, **kw)

    def _build(self, on_delete):
        p = dict(padx=3, pady=5)
        self._lbl(f"#{self._num:02d}", width=3).pack(side="left", **p)
        self._lbl("Src:", width=4).pack(side="left", **p)
        self._cb(self._src).pack(side="left", **p)
        self._lbl("→", fg=ACCENT).pack(side="left", **p)
        self._lbl("Tgt:", width=4).pack(side="left", **p)
        self._cb(self._tgt).pack(side="left", **p)
        self._freq_lbl = self._lbl(
            f"({BAND_FREQ['L18']}→{BAND_FREQ['TDD1']})", fg=ACCT2)
        self._freq_lbl.pack(side="left", padx=(2, 10), pady=5)
        self._lbl("cacHeadroom", width=11).pack(side="left", **p)
        self._ent(self._head, width=6).pack(side="left", **p)
        self._lbl("deltaCac", width=9).pack(side="left", **p)
        self._ent(self._delta, width=6).pack(side="left", **p)
        self._lbl("maxCacThreshold", width=15).pack(side="left", **p)
        self._ent(self._maxc, width=6).pack(side="left", **p)
        tk.Button(self, text="✕", font=MONO_S,
                  bg=PANEL2, fg=RED,
                  activebackground=PANEL2, activeforeground=RED,
                  relief="flat", bd=0, cursor="hand2",
                  command=on_delete).pack(side="right", padx=6)

    def _update_freq(self, _=None):
        s = BAND_FREQ.get(self._src.get(), "?")
        t = BAND_FREQ.get(self._tgt.get(), "?")
        self._freq_lbl.configure(text=f"({s}→{t})")

    def get(self):
        return {
            "source":          self._src.get(),
            "target":          self._tgt.get(),
            "cacHeadroom":     self._head.get().strip(),
            "deltaCac":        self._delta.get().strip(),
            "maxCacThreshold": self._maxc.get().strip(),
        }

    def set_values(self, src, tgt, headroom="5", delta="0", maxcac="100"):
        self._src.set(src)
        self._tgt.set(tgt)
        self._head.set(headroom)
        self._delta.set(delta)
        self._maxc.set(maxcac)
        self._update_freq()


# ─────────────────────────────────────────────────────────────────────────────
# Main application window
# ─────────────────────────────────────────────────────────────────────────────

class AutoAMLE(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("AutoAMLE  |  " + AUTHOR)
        self.state("zoomed")
        self.minsize(1050, 700)
        self.configure(bg=BG)
        self._mdb_path    = tk.StringVar()
        self._out_dir     = tk.StringVar(
            value=os.path.join(os.path.expanduser("~"), "Desktop"))
        self._band_rows: list = []
        self._row_counter = 0
        self._build_ui()
        self._add_band_row()

    def _build_ui(self):
        self._style()

        # ── header ──────────────────────────────────────────────────────────
        hdr = tk.Frame(self, bg=BG, pady=10)
        hdr.pack(fill="x", padx=20)
        tk.Label(hdr, text="🚀 AutoAMLE",
                 font=("Consolas", 15, "bold"),
                 fg=ACCENT, bg=BG).pack(side="left")
        tk.Label(hdr,
                 text="  –  Automating AMLE Profiles, LNREL Relations, and actAmle Activation",
                 font=MONO_S, fg=FG2, bg=BG).pack(side="left", pady=3)
        tk.Label(hdr,
                 text=AUTHOR,
                 font=("Consolas", 11, "bold"),
                 fg=GREEN, bg=BG).pack(side="right", padx=4)
        tk.Frame(self, bg=ACCENT, height=1).pack(fill="x", padx=20)

        # ── two-column body ──────────────────────────────────────────────────
        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True, padx=20, pady=8)

        left  = tk.Frame(body, bg=BG, width=390)
        right = tk.Frame(body, bg=BG)
        left.pack(side="left", fill="both", expand=False, padx=(0, 8))
        left.pack_propagate(False)
        right.pack(side="left", fill="both", expand=True)

        # ── LEFT ─────────────────────────────────────────────────────────────

        self._section(left, "1 · MDB DATABASE FILE")
        p = self._panel(left)
        self._entry_row(p, "MDB:", self._mdb_path,
                        btn="Browse…", cmd=self._browse_mdb)

        self._section(left, "2 · SITES  (one mrbtsId per line)")
        p = self._panel(left)
        self._sites_box = scrolledtext.ScrolledText(
            p, height=6, bg="#090d11", fg=FG,
            insertbackground=ACCENT, font=MONO,
            relief="flat", bd=0,
            highlightthickness=1, highlightbackground=BORDER,
            highlightcolor=ACCENT)
        self._sites_box.pack(fill="both", padx=5, pady=5)
        tk.Label(p, text="  e.g.  20027\n         20028",
                 font=("Consolas", 8), fg=FG2, bg=PANEL
                 ).pack(anchor="w", padx=5, pady=(0, 4))

        self._section(left, "3 · OUTPUT FOLDER")
        p = self._panel(left)
        self._entry_row(p, "Save:", self._out_dir,
                        btn="Browse…", cmd=self._browse_out)

        # ── Band Reference — updated freqs ───────────────────────────────────
        self._section(left, "BAND REFERENCE")
        p = self._panel(left)
        ref = (
            "  Band    Cells (lnCelId)       Freq\n"
            "  ─────  ───────────────────  ──────\n"
            "  L18    1-9  ,  99-199         1675\n"
            "  L9     10-19                  3622\n"
            "  L21    20-29,  200-299          224\n"
            "  TDD1   30-39                 40392\n"
            "  TDD2   40-49                 40590\n"
        )
        tk.Label(p, text=ref, font=("Consolas", 9),
                 fg=FG2, bg=PANEL, justify="left"
                 ).pack(anchor="w", padx=6, pady=4)

        self._section(left, "CSV SCOPE")
        p = self._panel(left)
        note = (
            "  AMLEPR   source-band cells only\n"
            "           (per mapping row)\n\n"
            "  LNREL    ALL cells in site\n"
            "           intra-site relations only\n\n"
            "  LC       ALL cells in site\n"
            "           no band filter"
        )
        tk.Label(p, text=note, font=("Consolas", 9),
                 fg=GREEN, bg=PANEL, justify="left"
                 ).pack(anchor="w", padx=6, pady=6)

        tk.Frame(left, bg=BG, height=6).pack()
        self._run_btn = tk.Button(
            left, text="▶  GENERATE CSVs",
            font=("Consolas", 12, "bold"),
            bg=ACCENT, fg=BG,
            activebackground=ACCT2, activeforeground=BG,
            relief="flat", bd=0, cursor="hand2", pady=10,
            command=self._run)
        self._run_btn.pack(fill="x")
        tk.Frame(left, bg=BG, height=4).pack()
        tk.Button(left, text="Clear Log",
                  font=MONO_S, bg=PANEL, fg=FG2,
                  activebackground=BORDER, activeforeground=FG,
                  relief="flat", bd=0, cursor="hand2", pady=4,
                  command=self._clear_log).pack(fill="x")

        tk.Label(left,
                 text=AUTHOR,
                 font=("Consolas", 9, "bold"),
                 fg=GREEN, bg=BG
                 ).pack(pady=(10, 0))

        # ── RIGHT ────────────────────────────────────────────────────────────

        self._section(right, "4 · BAND MAPPINGS  — Source → Target + CAC Parameters")
        tbl_frame = self._panel(right, expand=False)

        hdr2 = tk.Frame(tbl_frame, bg=PANEL)
        hdr2.pack(fill="x", padx=4, pady=(4, 0))
        for text, w in [("#",3),("Src",4),("",2),("Tgt",4),("Freqs",11),
                        ("cacHeadroom",12),("deltaCac",10),
                        ("maxCacThreshold",16),("",3)]:
            tk.Label(hdr2, text=text, font=("Consolas",8,"bold"),
                     fg=ACCENT, bg=PANEL, width=w, anchor="w"
                     ).pack(side="left", padx=3)

        cf = tk.Frame(tbl_frame, bg=PANEL)
        cf.pack(fill="both", expand=True, padx=4, pady=4)
        self._canvas = tk.Canvas(cf, bg=PANEL, highlightthickness=0, height=280)
        sb = ttk.Scrollbar(cf, orient="vertical", command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self._canvas.pack(side="left", fill="both", expand=True)
        self._rows_frame = tk.Frame(self._canvas, bg=PANEL)
        self._cwin = self._canvas.create_window(
            (0, 0), window=self._rows_frame, anchor="nw")
        self._rows_frame.bind("<Configure>",
            lambda e: self._canvas.configure(
                scrollregion=self._canvas.bbox("all")))
        self._canvas.bind("<Configure>",
            lambda e: self._canvas.itemconfig(self._cwin, width=e.width))

        btn_row = tk.Frame(tbl_frame, bg=PANEL)
        btn_row.pack(fill="x", padx=4, pady=4)
        tk.Button(btn_row, text="＋ Add Row",
                  font=MONO_S, bg=GREEN, fg=BG,
                  activebackground="#2ea043", activeforeground=BG,
                  relief="flat", bd=0, cursor="hand2", padx=10, pady=4,
                  command=self._add_band_row).pack(side="left")
        tk.Button(btn_row, text="Clear All",
                  font=MONO_S, bg=PANEL2, fg=FG2,
                  activebackground=BORDER, activeforeground=FG,
                  relief="flat", bd=0, cursor="hand2", padx=10, pady=4,
                  command=self._clear_rows).pack(side="left", padx=6)

        tk.Label(btn_row, text="|  Presets:",
                 font=MONO_S, fg=FG2, bg=PANEL).pack(side="left", padx=(10, 4))
        for label, cmd in [
            ("L18→All",  self._preset_l18_all),
            ("L9→All",   self._preset_l9_all),
            ("L21→All",  self._preset_l21_all),
            ("All→TDD1", self._preset_all_tdd1),
            ("All→TDD2", self._preset_all_tdd2),
        ]:
            tk.Button(btn_row, text=label,
                      font=MONO_S, bg=BORDER, fg=FG,
                      activebackground=ACCT2, activeforeground=BG,
                      relief="flat", bd=0, cursor="hand2", padx=6, pady=4,
                      command=cmd).pack(side="left", padx=2)

        self._section(right, "LOG OUTPUT")
        lp = self._panel(right, expand=True)
        self._log_box = scrolledtext.ScrolledText(
            lp, bg="#090d11", fg="#00ff88",
            insertbackground=ACCENT, font=("Consolas", 9),
            relief="flat", bd=0,
            highlightthickness=1, highlightbackground=BORDER,
            state="disabled")
        self._log_box.pack(fill="both", expand=True, padx=4, pady=4)

    def _style(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure("TCombobox",
                    fieldbackground=PANEL2, background=PANEL2,
                    foreground=FG, bordercolor=BORDER,
                    arrowcolor=ACCENT, selectbackground=BORDER)
        s.map("TCombobox",
              fieldbackground=[("readonly", PANEL2)],
              selectbackground=[("readonly", PANEL2)],
              foreground=[("readonly", FG)])
        s.configure("Vertical.TScrollbar",
                    background=BORDER, troughcolor=PANEL,
                    arrowcolor=FG2, bordercolor=PANEL)

    def _section(self, parent, title):
        f = tk.Frame(parent, bg=BG, pady=3)
        f.pack(fill="x")
        tk.Label(f, text=f"  {title}",
                 font=("Consolas", 9, "bold"), fg=ACCENT, bg=BG
                 ).pack(anchor="w")
        tk.Frame(f, bg=ACCENT, height=1).pack(fill="x")

    def _panel(self, parent, expand=False):
        p = tk.Frame(parent, bg=PANEL,
                     highlightthickness=1, highlightbackground=BORDER)
        p.pack(fill="both", expand=expand, pady=(2, 6))
        return p

    def _entry_row(self, parent, label, var, btn=None, cmd=None):
        row = tk.Frame(parent, bg=PANEL)
        row.pack(fill="x", padx=6, pady=5)
        tk.Label(row, text=label, font=MONO, fg=FG2,
                 bg=PANEL, width=5, anchor="w").pack(side="left")
        tk.Entry(row, textvariable=var,
                 bg="#090d11", fg=FG, insertbackground=ACCENT,
                 relief="flat", bd=0,
                 highlightthickness=1, highlightbackground=BORDER,
                 highlightcolor=ACCENT, font=MONO
                 ).pack(side="left", fill="x", expand=True)
        if btn:
            tk.Button(row, text=btn, font=MONO_S,
                      bg=BORDER, fg=FG,
                      activebackground=ACCENT, activeforeground=BG,
                      relief="flat", bd=0, cursor="hand2", padx=8,
                      command=cmd).pack(side="left", padx=(4, 0))

    def _add_band_row(self, src=None, tgt=None,
                      headroom="5", delta="0", maxcac="100"):
        self._row_counter += 1
        row = BandRow(self._rows_frame, self._row_counter,
                      on_delete=lambda r=None: self._delete_row(row))
        row.pack(fill="x", pady=2, padx=2)
        if src:
            row.set_values(src, tgt, headroom, delta, maxcac)
        self._band_rows.append(row)
        self._canvas.update_idletasks()
        self._canvas.yview_moveto(1.0)

    def _delete_row(self, row):
        if len(self._band_rows) <= 1:
            messagebox.showinfo("Info", "At least one mapping row is required.")
            return
        row.destroy()
        self._band_rows.remove(row)

    def _clear_rows(self):
        if not messagebox.askyesno("Clear", "Remove all band mapping rows?"):
            return
        for r in self._band_rows:
            r.destroy()
        self._band_rows.clear()
        self._row_counter = 0
        self._add_band_row()

    def _clear_log(self):
        self._log_box.configure(state="normal")
        self._log_box.delete("1.0", "end")
        self._log_box.configure(state="disabled")

    def _reset_rows(self):
        for r in self._band_rows:
            r.destroy()
        self._band_rows.clear()
        self._row_counter = 0

    def _preset_l18_all(self):
        self._reset_rows()
        [self._add_band_row("L18", t) for t in BANDS if t != "L18"]

    def _preset_l9_all(self):
        self._reset_rows()
        [self._add_band_row("L9", t) for t in BANDS if t != "L9"]

    def _preset_l21_all(self):
        self._reset_rows()
        [self._add_band_row("L21", t) for t in BANDS if t != "L21"]

    def _preset_all_tdd1(self):
        self._reset_rows()
        [self._add_band_row(s, "TDD1") for s in BANDS if s != "TDD1"]

    def _preset_all_tdd2(self):
        self._reset_rows()
        [self._add_band_row(s, "TDD2") for s in BANDS if s != "TDD2"]

    def _browse_mdb(self):
        p = filedialog.askopenfilename(
            title="Select MDB Dump",
            filetypes=[("Access Database", "*.mdb *.accdb"),
                       ("All Files", "*.*")])
        if p:
            self._mdb_path.set(p)

    def _browse_out(self):
        p = filedialog.askdirectory(title="Select Output Folder")
        if p:
            self._out_dir.set(p)

    def _log(self, msg):
        self._log_box.configure(state="normal")
        self._log_box.insert("end", msg + "\n")
        self._log_box.see("end")
        self._log_box.configure(state="disabled")
        self.update_idletasks()

    def _run(self):
        mdb = self._mdb_path.get().strip()
        if not mdb or not os.path.isfile(mdb):
            messagebox.showerror("Error", "Please select a valid MDB file.")
            return

        raw = self._sites_box.get("1.0", "end").strip()
        ids = [x.strip() for x in raw.splitlines() if x.strip()]
        if not ids:
            messagebox.showerror("Error", "Enter at least one mrbtsId.")
            return
        for s in ids:
            if not s.isdigit():
                messagebox.showerror("Error", f"Invalid mrbtsId: '{s}'")
                return

        if not self._band_rows:
            messagebox.showerror("Error", "Add at least one band mapping row.")
            return

        band_rows = []
        for i, r in enumerate(self._band_rows, 1):
            d = r.get()
            for field in ("cacHeadroom", "deltaCac", "maxCacThreshold"):
                if not d[field].lstrip("-").isdigit():
                    messagebox.showerror(
                        "Error",
                        f"Row {i}: '{field}' must be an integer "
                        f"(got '{d[field]}').")
                    return
            band_rows.append({
                "source":          d["source"],
                "target":          d["target"],
                "cacHeadroom":     int(d["cacHeadroom"]),
                "deltaCac":        int(d["deltaCac"]),
                "maxCacThreshold": int(d["maxCacThreshold"]),
            })

        out_dir = self._out_dir.get().strip()
        os.makedirs(out_dir, exist_ok=True)

        self._run_btn.configure(state="disabled", text="⏳  Processing…")

        def worker():
            try:
                ok = process(mdb, ids, band_rows, out_dir, self._log)
                if ok:
                    self.after(0, lambda: messagebox.showinfo(
                        "Done ✔", f"3 CSV files saved to:\n{out_dir}"))
                else:
                    self.after(0, lambda: messagebox.showwarning(
                        "Warning",
                        "Finished with warnings — please check the log."))
            except Exception as ex:
                import traceback
                self._log(f"\n[FATAL] {ex}\n{traceback.format_exc()}")
                self.after(0, lambda: messagebox.showerror("Error", str(ex)))
            finally:
                self.after(0, lambda: self._run_btn.configure(
                    state="normal", text="▶  GENERATE CSVs"))

        threading.Thread(target=worker, daemon=True).start()


# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = AutoAMLE()
    app.mainloop()
