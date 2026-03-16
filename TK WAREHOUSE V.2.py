import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import re
import math
import os
from datetime import datetime
import tempfile
import subprocess
import platform


# ----------------------------
# Helpers
# ----------------------------
def normalize_part(part) -> str:
    if pd.isna(part):
        return ""

    s = str(part).strip()

    # Remove only spreadsheet-added .0
    # Example: 001234.0 -> 001234
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]

    return s


def parse_qty(q):
    """
    Parse quantity:
    - 6 -> 6
    - "6 pcs" -> 6
    - "1 box" -> 1
    - "ALL" -> token
    """
    if pd.isna(q):
        return None, "missing"

    if isinstance(q, (int, float)) and not (isinstance(q, float) and math.isnan(q)):
        return int(q), "number"

    s = str(q).strip().lower()
    if s in ("all", "everything"):
        return "ALL", "all"

    m = re.search(r"-?\d+", s)
    if m:
        return int(m.group(0)), "number_in_text"

    return None, "unparsed"


def today_mdy2() -> str:
    s = datetime.now().strftime("%m/%d/%y")
    s = s.lstrip("0").replace("/0", "/")
    return s


def stamp_ymdhm() -> str:
    return datetime.now().strftime("%Y-%m-%d_%H%M")


def safe_default_filename(prefix: str, ext: str = "csv") -> str:
    return f"{prefix}_{stamp_ymdhm()}.{ext}"


# ----------------------------
# Smart paste parsing helpers
# ----------------------------
def looks_like_date(token: str) -> bool:
    token = str(token).strip()
    return bool(re.fullmatch(r"\d{1,2}/\d{1,2}/\d{2,4}", token))


def is_integer_token(token: str) -> bool:
    return bool(re.fullmatch(r"-?\d+", str(token).strip()))


def looks_like_part(token: str) -> bool:
    t = str(token).strip()

    if not t:
        return False

    if looks_like_date(t):
        return False

    banned = {
        "sam", "ashley", "joy", "shipper", "kl", "all",
        "warehouse", "cart", "desk", "aline", "mvp",
        "dnw1", "dnw2", "csr3", "cnr4", "ds", "fl", "v20",
        "c", "e", "details", "to"
    }
    if t.lower() in banned:
        return False

    if re.fullmatch(r"[A-Za-z]?\d{6,12}", t):
        return True

    return False


def clean_cell_value(v):
    if pd.isna(v):
        return ""
    return str(v).strip()


def split_pasted_line(line: str):
    raw = str(line).rstrip("\n").strip()
    if not raw:
        return []

    if "\t" in raw:
        return [c.strip() for c in raw.split("\t")]

    cols = re.split(r"\s{2,}", raw)
    if len(cols) > 1:
        return [c.strip() for c in cols]

    return [c.strip() for c in raw.split()]


def parse_block_record(block):
    """
    Expected layout:
    0  request date
    1  person
    2  job 1
    3  part number
    4  on hand qty
    5  room
    6  location
    7  pull qty
    8  due date
    9  job 2 / destination text
    """
    block = [clean_cell_value(x) for x in block if clean_cell_value(x) != ""]

    if len(block) < 9:
        return None

    request_date = block[0] if len(block) > 0 else ""
    person = block[1] if len(block) > 1 else ""
    job1 = block[2] if len(block) > 2 else ""
    part = normalize_part(block[3] if len(block) > 3 else "")
    room = block[5] if len(block) > 5 else ""
    loc = block[6] if len(block) > 6 else ""
    qty_raw = block[7] if len(block) > 7 else ""
    due_date = block[8] if len(block) > 8 else ""
    job2 = " ".join(block[9:]).strip() if len(block) > 9 else ""

    qty_val, _ = parse_qty(qty_raw)

    if part and qty_val is not None:
        return {
            "REQUEST_DATE": request_date,
            "PERSON": person,
            "Job": job1,
            "PART NUMBER": part,
            "ON_HAND_QTY": block[4] if len(block) > 4 else "",
            "RM": room,
            "Location": loc,
            "QTY PULLED": qty_val,
            "SHIP_DATE": due_date,
            "Job 2": job2
        }

    # Fallback smart scan
    part_candidates = []
    for i, c in enumerate(block):
        if looks_like_part(c):
            score = 0
            if re.fullmatch(r"\d{7,12}", c):
                score += 3
            if re.fullmatch(r"[A-Za-z]\d{6,12}", c):
                score += 3
            if i >= 2:
                score += 1
            part_candidates.append((score, i, c))

    if not part_candidates:
        return None

    part_candidates.sort(key=lambda x: (x[0], x[1]))
    _, best_idx, best_part = part_candidates[-1]
    part = normalize_part(best_part)

    job1 = block[best_idx - 1] if best_idx >= 1 else ""
    qty_val = None
    qty_idx = None

    for i, tok in enumerate(block):
        tl = str(tok).lower().strip()
        if tl == "all":
            qty_val = "ALL"
            qty_idx = i
            break
        if is_integer_token(tok):
            if i + 1 < len(block) and looks_like_date(block[i + 1]):
                qty_val = int(tok)
                qty_idx = i
                break

    room = block[qty_idx - 2] if qty_idx is not None and qty_idx - 2 >= 0 else ""
    loc = block[qty_idx - 1] if qty_idx is not None and qty_idx - 1 >= 0 else ""
    due_date = block[qty_idx + 1] if qty_idx is not None and qty_idx + 1 < len(block) else ""
    job2 = block[qty_idx + 2] if qty_idx is not None and qty_idx + 2 < len(block) else ""

    if not part or qty_val is None:
        return None

    return {
        "REQUEST_DATE": request_date,
        "PERSON": person,
        "Job": job1,
        "PART NUMBER": part,
        "ON_HAND_QTY": "",
        "RM": room,
        "Location": loc,
        "QTY PULLED": qty_val,
        "SHIP_DATE": due_date,
        "Job 2": job2
    }


def parse_pasted_pull_rows(raw_text: str):
    """
    Handles BOTH:
    1. tab-separated row paste from Sheets/Excel
    2. vertically stacked copied cells
    """
    output_cols = [
        "REQUEST_DATE", "PERSON", "Job", "PART NUMBER", "ON_HAND_QTY",
        "RM", "Location", "QTY PULLED", "SHIP_DATE", "Job 2"
    ]

    if not raw_text.strip():
        return pd.DataFrame(columns=output_cols)

    # First try direct row-based parsing from tabs (best for Sheets)
    parsed_rows = []
    for raw_line in raw_text.splitlines():
        line = raw_line.strip()
        if not line:
            continue

        low = line.lower()
        if "please bring today" in low or low.startswith("to:") or "pulls for" in low or low == "details":
            continue

        cols = split_pasted_line(line)
        cols = [clean_cell_value(c) for c in cols if clean_cell_value(c) != ""]

        # Exact 10-column layout from Sheets/email
        if len(cols) >= 10 and looks_like_date(cols[0]):
            record = {
                "REQUEST_DATE": cols[0],
                "PERSON": cols[1],
                "Job": cols[2],
                "PART NUMBER": normalize_part(cols[3]),
                "ON_HAND_QTY": cols[4],
                "RM": cols[5],
                "Location": cols[6],
                "QTY PULLED": parse_qty(cols[7])[0],
                "SHIP_DATE": cols[8],
                "Job 2": " ".join(cols[9:]).strip()
            }
            if record["PART NUMBER"] and record["QTY PULLED"] is not None:
                parsed_rows.append(record)
                continue

        # Fallback
        record = parse_block_record(cols)
        if record:
            parsed_rows.append(record)

    if parsed_rows:
        return pd.DataFrame(parsed_rows, columns=output_cols)

    # Vertical stacked mode
    raw_lines = [line.rstrip() for line in raw_text.splitlines()]
    lines = [clean_cell_value(x) for x in raw_lines if clean_cell_value(x) != ""]

    date_positions = [i for i, x in enumerate(lines) if looks_like_date(x)]

    if date_positions and looks_like_date(lines[0]):
        blocks = []
        current = []

        for line in lines:
            if looks_like_date(line) and current:
                blocks.append(current)
                current = [line]
            else:
                current.append(line)

        if current:
            blocks.append(current)

        parsed = []
        for blk in blocks:
            record = parse_block_record(blk)
            if record:
                parsed.append(record)

        if parsed:
            return pd.DataFrame(parsed, columns=output_cols)

    return pd.DataFrame(columns=output_cols)


# ----------------------------
# App
# ----------------------------
class WarehouseApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Tekmor Warehouse Tool")
        self.geometry("1350x780")

        # Data
        self.warehouse_df = None
        self.pull_df = None
        self.log_df = pd.DataFrame()
        self.last_warehouse_path = None
        self._last_shortages_df = pd.DataFrame()

        # Warehouse columns
        self.W_PART = "Part"
        self.W_QTY = "OH Now"
        self.W_DATE_MAIN = "Date"
        self.W_DATE = "Last Updated"
        self.W_INOUT = "In/Out"

        # Auto rename old comments column
        self.AUTO_RENAME_COMMENTS_TO_INOUT = True
        self.OLD_COMMENTS_NAME = "Comments"

        # Pull columns
        self.P_PART = "PART NUMBER"
        self.P_QTY = "QTY PULLED"

        # Zebra / pull optional columns
        self.P_JOB = "Job"
        self.P_JOB2 = "Job 2"
        self.P_RM = "RM"
        self.P_LOC = "Location"

        # Fallback positional mapping
        self.P_COL_JOB_IDX = 2
        self.P_COL_PART_IDX = 3
        self.P_COL_RM_IDX = 5
        self.P_COL_LOC_IDX = 6
        self.P_COL_QTY_IDX = 7
        self.P_COL_JOB2_IDX = 9

        # Zebra printer name
        self.ZEBRA_PRINTER_NAME = "Zebra_Technologies_ZTC_ZT230_200dpi_ZPL"

        # Safety options
        self.DRY_RUN_MODE = tk.BooleanVar(value=False)
        self.AUTO_BACKUP_BEFORE_APPLY = True

        self._build_menu()
        self._build_layout()

    # ---------- UI ----------
    def _build_menu(self):
        menubar = tk.Menu(self)

        file_menu = tk.Menu(menubar, tearoff=False)
        file_menu.add_command(label="Load Warehouse CSV...", command=self.load_warehouse)
        file_menu.add_command(label="Load Pull List CSV...", command=self.load_pull_list)
        file_menu.add_command(label="Paste Pull List...", command=self.open_paste_pull_list_window)
        file_menu.add_separator()
        file_menu.add_command(label="Save Updated Warehouse As...", command=self.save_updated_warehouse)
        file_menu.add_command(label="Export Log As...", command=self.export_log)
        file_menu.add_command(label="Export Shortages As...", command=self.export_shortages)
        file_menu.add_separator()
        file_menu.add_command(label="Export Zebra Tags (.zpl)...", command=self.export_zebra_tags)
        file_menu.add_command(label="Print Zebra Tags...", command=self.print_zebra_tags)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.destroy)
        menubar.add_cascade(label="File", menu=file_menu)

        inv_menu = tk.Menu(menubar, tearoff=False)
        inv_menu.add_command(label="Refresh Inventory View", command=self.refresh_inventory_view)
        inv_menu.add_command(label="Show Low Stock (<= 0)", command=self.filter_low_stock)
        inv_menu.add_command(label="Clear Filters", command=self.clear_filters)
        inv_menu.add_separator()
        inv_menu.add_command(label="Manual Receive (+)", command=lambda: self.open_manual_adjust_window("receive"))
        inv_menu.add_command(label="Manual Send (-)", command=lambda: self.open_manual_adjust_window("send"))
        menubar.add_cascade(label="Inventory", menu=inv_menu)

        pull_menu = tk.Menu(menubar, tearoff=False)
        pull_menu.add_command(label="Preview Pull List", command=self.preview_pull_list)
        pull_menu.add_command(label="Paste Pull List", command=self.open_paste_pull_list_window)
        pull_menu.add_command(label="Apply Pull List to Inventory", command=self.apply_pull_list)
        menubar.add_cascade(label="Pull List", menu=pull_menu)

        tools_menu = tk.Menu(menubar, tearoff=False)
        tools_menu.add_command(label="Settings...", command=self.open_settings)
        tools_menu.add_separator()
        tools_menu.add_checkbutton(label="Dry Run / Test Mode (no changes)", variable=self.DRY_RUN_MODE)
        menubar.add_cascade(label="Tools", menu=tools_menu)

        help_menu = tk.Menu(menubar, tearoff=False)
        help_menu.add_command(label="About", command=self.about)
        menubar.add_cascade(label="Help", menu=help_menu)

        self.config(menu=menubar)

    def _build_layout(self):
        top = ttk.Frame(self, padding=10)
        top.pack(side=tk.TOP, fill=tk.X)

        ttk.Label(top, text="Part Search:").pack(side=tk.LEFT)
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(top, textvariable=self.search_var, width=25)
        self.search_entry.pack(side=tk.LEFT, padx=8)
        self.search_entry.bind("<Return>", lambda e: self.search_part())

        ttk.Button(top, text="Search", command=self.search_part).pack(side=tk.LEFT)
        ttk.Button(top, text="Clear", command=self.clear_search).pack(side=tk.LEFT, padx=6)

        ttk.Separator(top, orient="vertical").pack(side=tk.LEFT, fill=tk.Y, padx=10)

        ttk.Button(top, text="Load Warehouse", command=self.load_warehouse).pack(side=tk.LEFT)
        ttk.Button(top, text="Load Pull List", command=self.load_pull_list).pack(side=tk.LEFT, padx=6)
        ttk.Button(top, text="Paste Pull List", command=self.open_paste_pull_list_window).pack(side=tk.LEFT, padx=6)
        ttk.Button(top, text="Apply Pull", command=self.apply_pull_list).pack(side=tk.LEFT, padx=6)

        self.status_var = tk.StringVar(value="Load a warehouse CSV to begin.")
        ttk.Label(top, textvariable=self.status_var).pack(side=tk.RIGHT)

        main = ttk.Frame(self, padding=(10, 0, 10, 10))
        main.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        left = ttk.Frame(main, width=450)
        left.pack(side=tk.LEFT, fill=tk.Y)

        ttk.Label(left, text="Part Details", font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(0, 8))
        self.details_text = tk.Text(left, height=12, wrap="word")
        self.details_text.pack(fill=tk.X, pady=(0, 10))
        self.details_text.configure(state="disabled")

        ttk.Label(left, text="Shipment Summary (after pull)", font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(10, 8))
        self.shipment_text = tk.Text(left, height=14, wrap="word")
        self.shipment_text.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        self.shipment_text.configure(state="disabled")

        ttk.Label(left, text="Pull List Preview", font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(0, 8))
        self.pull_preview = tk.Text(left, height=10, wrap="word")
        self.pull_preview.pack(fill=tk.BOTH, expand=True)
        self.pull_preview.configure(state="disabled")

        right = ttk.Frame(main)
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(12, 0))

        ttk.Label(right, text="Warehouse Inventory", font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(0, 8))

        table_frame = ttk.Frame(right)
        table_frame.pack(fill=tk.BOTH, expand=True)

        self.tree = ttk.Treeview(table_frame, show="headings")
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=vsb.set)

        hsb = ttk.Scrollbar(right, orient="horizontal", command=self.tree.xview)
        hsb.pack(fill=tk.X)
        self.tree.configure(xscrollcommand=hsb.set)

        self.tree.bind("<<TreeviewSelect>>", self.on_row_select)

    # ---------- Text setters ----------
    def _set_details(self, text):
        self.details_text.configure(state="normal")
        self.details_text.delete("1.0", tk.END)
        self.details_text.insert(tk.END, text)
        self.details_text.configure(state="disabled")

    def _set_pull_preview(self, text):
        self.pull_preview.configure(state="normal")
        self.pull_preview.delete("1.0", tk.END)
        self.pull_preview.insert(tk.END, text)
        self.pull_preview.configure(state="disabled")

    def _set_shipment(self, text):
        self.shipment_text.configure(state="normal")
        self.shipment_text.delete("1.0", tk.END)
        self.shipment_text.insert(tk.END, text)
        self.shipment_text.configure(state="disabled")

    # ---------- Data actions ----------
    def load_warehouse(self):
        path = filedialog.askopenfilename(
            title="Select Warehouse CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if not path:
            return
        try:
            df = pd.read_csv(path, dtype=str, keep_default_na=False)
        except Exception as e:
            messagebox.showerror("Load Error", f"Could not load warehouse file.\n\n{e}")
            return

        self.last_warehouse_path = path
        df.columns = [str(c).strip() for c in df.columns]

        if self.AUTO_RENAME_COMMENTS_TO_INOUT:
            if self.W_INOUT not in df.columns and self.OLD_COMMENTS_NAME in df.columns:
                df = df.rename(columns={self.OLD_COMMENTS_NAME: self.W_INOUT})

        if self.W_PART in df.columns:
            df[self.W_PART] = df[self.W_PART].apply(normalize_part)

        if self.W_QTY in df.columns:
            df[self.W_QTY] = pd.to_numeric(df[self.W_QTY], errors="coerce").fillna(0).astype(int)

        for col in (self.W_DATE_MAIN, self.W_DATE, self.W_INOUT):
            if col not in df.columns:
                df[col] = ""

        self.warehouse_df = df
        self._clear_tree()
        self.refresh_inventory_view()
        self.status_var.set(f"Warehouse loaded: {os.path.basename(path)}")

    def load_pull_list(self):
        path = filedialog.askopenfilename(
            title="Select Pull List CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if not path:
            return
        try:
            df = pd.read_csv(path, dtype=str, keep_default_na=False)
        except Exception as e:
            messagebox.showerror("Load Error", f"Could not load pull list file.\n\n{e}")
            return

        df.columns = [str(c).strip() for c in df.columns]

        if self.P_PART in df.columns:
            df[self.P_PART] = df[self.P_PART].apply(normalize_part)

        self.pull_df = df
        self.preview_pull_list()
        self.status_var.set(f"Pull list loaded: {os.path.basename(path)}")

    def open_paste_pull_list_window(self):
        win = tk.Toplevel(self)
        win.title("Paste Pull List")
        win.geometry("900x650")
        win.grab_set()

        frm = ttk.Frame(win, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            frm,
            text="Paste the full pull list below",
            font=("Segoe UI", 11, "bold")
        ).pack(anchor="w", pady=(0, 8))

        ttk.Label(
            frm,
            text="Expected columns: Date | Person | Job 1 | Part Number | On Hand Qty | Room | Location | Pull Qty | Due Date | Job 2",
        ).pack(anchor="w", pady=(0, 8))

        text_box = tk.Text(frm, wrap="none")
        text_box.pack(fill=tk.BOTH, expand=True)

        preview_label = ttk.Label(frm, text="")
        preview_label.pack(anchor="w", pady=(8, 0))

        btns = ttk.Frame(frm)
        btns.pack(fill=tk.X, pady=(10, 0))

        def load_pasted_data():
            raw = text_box.get("1.0", tk.END).strip()
            if not raw:
                messagebox.showinfo("No Data", "Paste a pull list first.")
                return

            try:
                df = parse_pasted_pull_rows(raw)

                if df.empty:
                    raise ValueError(
                        "No valid pull rows were detected.\n"
                        "Try copying the rows from the email/table again."
                    )

                df[self.P_PART] = df[self.P_PART].apply(normalize_part)

                self.pull_df = df
                self.preview_pull_list()

                preview_label.config(text=f"Loaded {len(df)} rows from pasted data.")
                self.status_var.set(f"Pasted pull list loaded ({len(df)} rows).")
                messagebox.showinfo("Loaded", f"Pasted pull list loaded successfully.\nRows: {len(df)}")
                win.destroy()

            except Exception as e:
                messagebox.showerror("Paste Error", f"Could not parse pasted pull list.\n\n{e}")

        ttk.Button(btns, text="Load Pasted Pull List", command=load_pasted_data).pack(side=tk.RIGHT)
        ttk.Button(btns, text="Cancel", command=win.destroy).pack(side=tk.RIGHT, padx=(0, 8))

    def refresh_inventory_view(self, df_to_show=None):
        if self.warehouse_df is None:
            messagebox.showinfo("No Data", "Load a warehouse file first.")
            return
        df = self.warehouse_df if df_to_show is None else df_to_show
        self._populate_tree(df)

    # ---------- Lookups ----------
    def _get_part_index(self, part: str):
        if self.warehouse_df is None:
            return None
        if self.W_PART not in self.warehouse_df.columns:
            return None

        part = normalize_part(part)
        matches = self.warehouse_df[self.W_PART] == part
        matched = self.warehouse_df[matches]
        if matched.empty:
            return None
        return int(matched.index[0])

    def _get_selected_part(self):
        sel = self.tree.selection()
        if not sel:
            return ""
        item = sel[0]
        vals = self.tree.item(item, "values")
        cols = list(self.tree["columns"])
        if self.W_PART in cols:
            idx = cols.index(self.W_PART)
            if idx < len(vals):
                return normalize_part(vals[idx])
        return ""

    # ---------- Pull planning ----------
    def build_pull_plan(self):
        wh = self.warehouse_df
        pull = self.pull_df

        index_by_part = {}
        for idx, part in wh[self.W_PART].items():
            if part and part not in index_by_part:
                index_by_part[part] = idx

        run_ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        run_date = today_mdy2()

        plan_rows = []
        log_rows = []
        shortage_rows = []

        for _, row in pull.iterrows():
            part = normalize_part(row.get(self.P_PART, ""))
            qty_raw = row.get(self.P_QTY, None)

            if not part:
                log_rows.append({
                    "timestamp": run_ts,
                    "category": "skipped",
                    "part": "",
                    "qty_raw": qty_raw,
                    "reason": "missing_part"
                })
                continue

            qty, qty_type = parse_qty(qty_raw)
            if qty is None:
                log_rows.append({
                    "timestamp": run_ts,
                    "category": "skipped",
                    "part": part,
                    "qty_raw": qty_raw,
                    "reason": f"qty_{qty_type}"
                })
                continue

            if part not in index_by_part:
                log_rows.append({
                    "timestamp": run_ts,
                    "category": "not_found",
                    "part": part,
                    "qty_raw": qty_raw
                })
                continue

            w_idx = index_by_part[part]
            before = int(wh.at[w_idx, self.W_QTY])
            requested = before if qty == "ALL" else int(qty)
            sent = min(before, requested)
            after = before - sent

            if requested > before:
                short = requested - sent
                log_rows.append({
                    "timestamp": run_ts,
                    "category": "partial_fill",
                    "part": part,
                    "requested": requested,
                    "sent": sent,
                    "available": before
                })
                shortage_rows.append({
                    "part": part,
                    "requested": requested,
                    "available": before,
                    "sent": sent,
                    "short": short
                })

            existing_main_date = wh.at[w_idx, self.W_DATE_MAIN] if self.W_DATE_MAIN in wh.columns else ""
            existing_last_updated = wh.at[w_idx, self.W_DATE] if self.W_DATE in wh.columns else ""
            existing_inout = wh.at[w_idx, self.W_INOUT] if self.W_INOUT in wh.columns else ""

            effective_date_main = run_date if sent > 0 else existing_main_date
            effective_last_updated = run_date if sent > 0 else existing_last_updated
            effective_inout = -sent if sent > 0 else existing_inout

            plan_rows.append({
                "part": part,
                "requested": requested,
                "sent": sent,
                "before": before,
                "after": after,
                "date_main": effective_date_main,
                "last_updated": effective_last_updated,
                "date": run_date,
                "inout": effective_inout,
                "_w_idx": w_idx
            })

        plan_df = pd.DataFrame(plan_rows)
        log_df = pd.DataFrame(log_rows)
        shortages_df = pd.DataFrame(shortage_rows)

        if not plan_df.empty:
            plan_df["_order"] = range(len(plan_df))
            final_rows = []
            current_after = {}

            for _, r in plan_df.sort_values("_order").iterrows():
                p = r["part"]
                base_before = r["before"] if p not in current_after else current_after[p]
                requested = int(r["requested"])
                sent = min(base_before, requested)
                after = base_before - sent
                current_after[p] = after

                w_idx = int(r["_w_idx"])
                existing_main_date = wh.at[w_idx, self.W_DATE_MAIN] if self.W_DATE_MAIN in wh.columns else ""
                existing_last_updated = wh.at[w_idx, self.W_DATE] if self.W_DATE in wh.columns else ""
                existing_inout = wh.at[w_idx, self.W_INOUT] if self.W_INOUT in wh.columns else ""

                final_rows.append({
                    "part": p,
                    "requested": requested,
                    "sent": sent,
                    "before": base_before,
                    "after": after,
                    "date_main": today_mdy2() if sent > 0 else existing_main_date,
                    "last_updated": today_mdy2() if sent > 0 else existing_last_updated,
                    "date": today_mdy2(),
                    "inout": -sent if sent > 0 else existing_inout,
                    "_w_idx": w_idx
                })

            plan_df = pd.DataFrame(final_rows)

            plan_df = plan_df.groupby("part", as_index=False).agg(
                requested=("requested", "sum"),
                sent=("sent", "sum"),
                before=("before", "first"),
                after=("after", "last"),
                date_main=("date_main", "last"),
                last_updated=("last_updated", "last"),
                date=("date", "first"),
                inout=("inout", "last"),
                _w_idx=("_w_idx", "first"),
            )

        return plan_df, log_df, shortages_df

    # ---------- Confirmation windows ----------
    def confirm_pull_window(self, plan_df: pd.DataFrame) -> bool:
        if plan_df is None or plan_df.empty:
            messagebox.showinfo("Nothing to Apply", "No valid parts to apply (everything skipped/not found).")
            return False

        win = tk.Toplevel(self)
        win.title("Confirm Pull — Review Before Applying")
        win.geometry("1040x560")
        win.grab_set()

        ttk.Label(
            win,
            text="Review the pull below. Click APPLY to update inventory, or CANCEL to stop.",
            font=("Segoe UI", 11, "bold")
        ).pack(anchor="w", padx=12, pady=(12, 8))

        cols = ["part", "requested", "sent", "before", "after", "date_main", "inout"]
        tree = ttk.Treeview(win, columns=cols, show="headings")
        tree.pack(fill=tk.BOTH, expand=True, padx=12, pady=8)

        headings = {
            "part": "PART",
            "requested": "REQUESTED",
            "sent": "SENT",
            "before": "BEFORE",
            "after": "AFTER",
            "date_main": "DATE",
            "inout": "IN/OUT"
        }

        for c in cols:
            tree.heading(c, text=headings.get(c, c.upper()))
            tree.column(c, width=130, stretch=True)

        for _, r in plan_df.iterrows():
            tree.insert("", "end", values=[r[c] for c in cols])

        totals = {
            "requested": int(plan_df["requested"].sum()),
            "sent": int(plan_df["sent"].sum()),
        }
        ttk.Label(
            win,
            text=f"TOTALS — Requested: {totals['requested']}   |   Sent: {totals['sent']}",
        ).pack(anchor="w", padx=12, pady=(0, 10))

        btn_frame = ttk.Frame(win)
        btn_frame.pack(fill=tk.X, padx=12, pady=(0, 12))

        result = {"ok": False}

        def on_apply():
            result["ok"] = True
            win.destroy()

        def on_cancel():
            result["ok"] = False
            win.destroy()

        ttk.Button(btn_frame, text="APPLY PULL", command=on_apply).pack(side=tk.RIGHT, padx=(8, 0))
        ttk.Button(btn_frame, text="Cancel", command=on_cancel).pack(side=tk.RIGHT)

        self.wait_window(win)
        return result["ok"]

    def confirm_manual_adjust_window(self, part, mode, qty, before, delta, after, date_display):
        win = tk.Toplevel(self)
        win.title("Confirm Manual Update")
        win.geometry("520x300")
        win.grab_set()

        action_name = "Receive (+)" if mode == "receive" else "Send (-)"

        box = ttk.Frame(win, padding=16)
        box.pack(fill=tk.BOTH, expand=True)

        ttk.Label(box, text=f"Manual {action_name}", font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(0, 10))

        lines = [
            f"Part: {part}",
            f"Entered Qty: {qty}",
            f"Before: {before}",
            f"Applied Change: {delta:+d}",
            f"After: {after}",
            f"Date: {date_display}",
        ]

        if mode == "send" and qty > before:
            lines.append("")
            lines.append(f"Requested {qty}, but only {before} available.")
            lines.append(f"System will send {before} and zero the part out.")

        if mode == "send" and delta == 0:
            lines.append("")
            lines.append("Nothing will be sent. Row will stay unchanged.")

        for line in lines:
            ttk.Label(box, text=line).pack(anchor="w")

        result = {"ok": False}

        btns = ttk.Frame(box)
        btns.pack(fill=tk.X, pady=(18, 0))

        def apply_now():
            result["ok"] = True
            win.destroy()

        def cancel_now():
            result["ok"] = False
            win.destroy()

        ttk.Button(btns, text="Apply", command=apply_now).pack(side=tk.RIGHT, padx=(8, 0))
        ttk.Button(btns, text="Cancel", command=cancel_now).pack(side=tk.RIGHT)

        self.wait_window(win)
        return result["ok"]

    # ---------- Safety ----------
    def _auto_backup(self):
        if not self.AUTO_BACKUP_BEFORE_APPLY:
            return None
        if self.warehouse_df is None:
            return None

        folder = os.path.dirname(self.last_warehouse_path) if self.last_warehouse_path else os.getcwd()
        backup_path = os.path.join(folder, safe_default_filename("BACKUP_before_change"))
        try:
            self.warehouse_df.to_csv(backup_path, index=False)
            return backup_path
        except Exception:
            return None

    # ---------- Pull apply ----------
    def apply_pull_list(self):
        if self.warehouse_df is None:
            messagebox.showinfo("Missing Warehouse", "Load a warehouse file first.")
            return
        if self.pull_df is None:
            messagebox.showinfo("Missing Pull List", "Load a pull list first.")
            return

        wh = self.warehouse_df
        pull = self.pull_df

        for col in [self.W_PART, self.W_QTY]:
            if col not in wh.columns:
                messagebox.showerror("Config Error", f"Warehouse missing column: {col} (Tools > Settings).")
                return
        for col in [self.P_PART, self.P_QTY]:
            if col not in pull.columns:
                messagebox.showerror("Config Error", f"Pull list missing column: {col} (Tools > Settings).")
                return

        for col in (self.W_DATE_MAIN, self.W_DATE, self.W_INOUT):
            if col not in wh.columns:
                wh[col] = ""

        plan_df, pre_log_df, shortages_df = self.build_pull_plan()

        ok = self.confirm_pull_window(plan_df)
        if not ok:
            self.status_var.set("Pull cancelled (no changes applied).")
            return

        if self.DRY_RUN_MODE.get():
            self._post_run_outputs(plan_df, pre_log_df, shortages_df)
            messagebox.showinfo("Dry Run Complete", "Dry Run mode is ON.\n\nNo inventory was changed.")
            return

        backup_path = self._auto_backup()

        run_ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        changes = []
        shipment_rows = []

        for _, r in plan_df.iterrows():
            part = r["part"]
            sent = int(r["sent"])
            before = int(r["before"])
            after = int(r["after"])
            w_idx = int(r["_w_idx"])

            if sent <= 0:
                continue

            run_date_main = r["date_main"]
            run_last_updated = r["last_updated"]

            wh.at[w_idx, self.W_QTY] = after
            wh.at[w_idx, self.W_DATE_MAIN] = run_date_main
            wh.at[w_idx, self.W_DATE] = run_last_updated
            wh.at[w_idx, self.W_INOUT] = str(-sent)

            changes.append({
                "timestamp": run_ts,
                "category": "changed",
                "part": part,
                "before": before,
                "pulled": sent,
                "after": after,
                "date": run_date_main,
                "inout": -sent
            })

            shipment_rows.append({
                "part": part,
                "sent": sent,
                "before": before,
                "after": after
            })

        batch_log = pd.concat([pre_log_df, pd.DataFrame(changes)], ignore_index=True)
        if self.log_df is None or self.log_df.empty:
            self.log_df = batch_log
        else:
            self.log_df = pd.concat([self.log_df, batch_log], ignore_index=True)

        self.refresh_inventory_view()
        self.preview_pull_list(summary=self._summarize_batch_log(batch_log))
        self._show_shipment_summary(shipment_rows)
        self._last_shortages_df = shortages_df

        msg = "✅ Pull applied successfully!"
        if backup_path:
            msg += f"\n\nBackup saved:\n{backup_path}"
        messagebox.showinfo("Pull Applied", msg)

    # ---------- Manual receive/send ----------
    def open_manual_adjust_window(self, mode="receive"):
        if self.warehouse_df is None:
            messagebox.showinfo("Missing Warehouse", "Load a warehouse file first.")
            return

        win = tk.Toplevel(self)
        win.title("Manual Receive (+)" if mode == "receive" else "Manual Send (-)")
        win.geometry("520x340")
        win.grab_set()

        frm = ttk.Frame(win, padding=16)
        frm.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            frm,
            text="Manual Receive (+)" if mode == "receive" else "Manual Send (-)",
            font=("Segoe UI", 12, "bold")
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 12))

        ttk.Label(frm, text="Part Number:").grid(row=1, column=0, sticky="w")
        part_var = tk.StringVar()

        selected_part = self._get_selected_part() or normalize_part(self.search_var.get())
        part_var.set(selected_part)

        part_entry = ttk.Entry(frm, textvariable=part_var, width=28)
        part_entry.grid(row=1, column=1, sticky="ew", pady=4)

        ttk.Label(frm, text="Quantity:").grid(row=2, column=0, sticky="w")
        qty_var = tk.StringVar()
        ttk.Entry(frm, textvariable=qty_var, width=28).grid(row=2, column=1, sticky="ew", pady=4)

        preview_text = tk.Text(frm, height=8, wrap="word")
        preview_text.grid(row=3, column=0, columnspan=2, sticky="nsew", pady=(12, 8))
        preview_text.configure(state="disabled")

        frm.columnconfigure(1, weight=1)
        frm.rowconfigure(3, weight=1)

        def set_preview(text):
            preview_text.configure(state="normal")
            preview_text.delete("1.0", tk.END)
            preview_text.insert(tk.END, text)
            preview_text.configure(state="disabled")

        def preview():
            part = normalize_part(part_var.get())
            qty_raw = qty_var.get().strip()

            if not part:
                set_preview("Enter a part number.")
                return None
            if not qty_raw:
                set_preview("Enter a quantity.")
                return None

            try:
                qty = int(qty_raw)
            except ValueError:
                set_preview("Quantity must be a whole number.")
                return None

            if qty < 0:
                set_preview("Quantity must be 0 or greater.")
                return None

            w_idx = self._get_part_index(part)
            if w_idx is None:
                set_preview(f"Part not found: {part}")
                return None

            before = int(self.warehouse_df.at[w_idx, self.W_QTY])
            existing_main_date = self.warehouse_df.at[w_idx, self.W_DATE_MAIN] if self.W_DATE_MAIN in self.warehouse_df.columns else ""
            existing_last_updated = self.warehouse_df.at[w_idx, self.W_DATE] if self.W_DATE in self.warehouse_df.columns else ""
            existing_inout = self.warehouse_df.at[w_idx, self.W_INOUT] if self.W_INOUT in self.warehouse_df.columns else ""

            today = today_mdy2()

            if mode == "receive":
                applied = qty
                after = before + applied
                date_main = today if applied > 0 else existing_main_date
                last_updated = today if applied > 0 else existing_last_updated
                inout = f"+{applied}" if applied > 0 else existing_inout
                action_label = "Receive"
            else:
                applied = min(before, qty)
                after = before - applied
                date_main = today if applied > 0 else existing_main_date
                last_updated = today if applied > 0 else existing_last_updated
                inout = -applied if applied > 0 else existing_inout
                action_label = "Send"

            lines = [
                f"Action: {action_label}",
                f"Part: {part}",
                f"Entered Qty: {qty}",
                f"Before: {before}",
                f"Applied Qty: {applied}",
                f"After: {after}",
                f"Date: {date_main}",
                f"In/Out: {inout}",
            ]

            if mode == "send" and qty > before:
                lines.append("")
                lines.append(f"Requested {qty}, but only {before} available.")
                lines.append(f"The system will send {applied} and stop at 0.")

            if mode == "send" and applied == 0:
                lines.append("")
                lines.append("Nothing is leaving inventory.")
                lines.append("The row will stay unchanged.")

            set_preview("\n".join(lines))

            return {
                "part": part,
                "entered_qty": qty,
                "applied_qty": applied,
                "before": before,
                "after": after,
                "date_main": date_main,
                "last_updated": last_updated,
                "inout": inout,
                "w_idx": w_idx
            }

        def apply_manual():
            result = preview()
            if result is None:
                return

            part = result["part"]
            qty = result["entered_qty"]
            applied = result["applied_qty"]
            before = result["before"]
            after = result["after"]
            date_main = result["date_main"]
            w_idx = result["w_idx"]

            ok = self.confirm_manual_adjust_window(
                part=part,
                mode=mode,
                qty=qty,
                before=before,
                delta=applied if mode == "receive" else -applied,
                after=after,
                date_display=date_main
            )
            if not ok:
                return

            if self.DRY_RUN_MODE.get():
                msg = "Dry Run mode is ON.\n\nNo inventory was changed."
                set_preview(f"{preview_text.get('1.0', tk.END).strip()}\n\n{msg}")
                self.status_var.set("Manual update previewed in Dry Run mode.")
                return

            backup_path = self._auto_backup()

            if mode == "send" and applied <= 0:
                messagebox.showinfo("No Change", "Nothing was sent, so the row was left unchanged.")
                if backup_path:
                    self.status_var.set(f"No change applied. Backup saved: {os.path.basename(backup_path)}")
                win.destroy()
                return

            if mode == "receive" and applied <= 0:
                messagebox.showinfo("No Change", "Quantity was 0, so the row was left unchanged.")
                if backup_path:
                    self.status_var.set(f"No change applied. Backup saved: {os.path.basename(backup_path)}")
                win.destroy()
                return

            self.warehouse_df.at[w_idx, self.W_QTY] = after
            self.warehouse_df.at[w_idx, self.W_DATE_MAIN] = result["date_main"]
            self.warehouse_df.at[w_idx, self.W_DATE] = result["last_updated"]
            self.warehouse_df.at[w_idx, self.W_INOUT] = result["inout"]

            run_ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            log_row = {
                "timestamp": run_ts,
                "category": "manual_receive" if mode == "receive" else "manual_send",
                "part": part,
                "before": before,
                "qty": applied,
                "after": after,
                "date": result["date_main"],
                "inout": result["inout"]
            }

            self.log_df = pd.concat([self.log_df, pd.DataFrame([log_row])], ignore_index=True)
            self.refresh_inventory_view()

            summary_lines = [
                f"Manual {'Receive' if mode == 'receive' else 'Send'} — Date: {result['date_main']}",
                "-" * 60,
                f"Part: {part}",
                f"Before: {before}",
                f"Qty: {applied}",
                f"After: {after}",
            ]
            self._set_shipment("\n".join(summary_lines))

            self.search_var.set(part)
            self.search_part()

            msg = "✅ Manual update applied!"
            if backup_path:
                msg += f"\n\nBackup saved:\n{backup_path}"
            messagebox.showinfo("Done", msg)
            win.destroy()

        ttk.Button(frm, text="Preview", command=preview).grid(row=4, column=0, sticky="w", pady=(6, 0))
        ttk.Button(frm, text="Apply", command=apply_manual).grid(row=4, column=1, sticky="e", pady=(6, 0))

        part_entry.focus_set()

    # ---------- Zebra tags ----------
    def _zpl_safe(self, value):
        if value is None:
            return ""
        s = str(value).strip()
        s = s.replace("^", "")
        s = s.replace("~", "")
        return s

    def _row_val_by_idx(self, row, idx):
        try:
            return row.iloc[idx]
        except Exception:
            return ""

    def _pull_value(self, row, named_col, fallback_idx):
        if named_col in row.index:
            return row.get(named_col, "")
        return self._row_val_by_idx(row, fallback_idx)

    def _combine_job_text(self, job1, job2):
        job1 = "" if pd.isna(job1) else str(job1).strip()
        job2 = "" if pd.isna(job2) else str(job2).strip()

        if job1 and job2:
            return f"{job1} {job2}"
        if job1:
            return job1
        if job2:
            return job2
        return ""

    def build_tag_rows(self):
        """
        Build ONE tag per pull-list row, using actual sent qty logic.
        Combines Job + Job 2 onto the tag.
        """
        if self.warehouse_df is None or self.pull_df is None:
            return []

        wh = self.warehouse_df
        pull = self.pull_df.copy()

        wh_index = {}
        for idx, part in wh[self.W_PART].items():
            if part and part not in wh_index:
                wh_index[part] = idx

        remaining = {}
        for idx, part in wh[self.W_PART].items():
            remaining[part] = int(wh.at[idx, self.W_QTY])

        tag_rows = []
        run_date = today_mdy2()

        for _, row in pull.iterrows():
            part = normalize_part(self._pull_value(row, self.P_PART, self.P_COL_PART_IDX))
            qty_raw = self._pull_value(row, self.P_QTY, self.P_COL_QTY_IDX)

            if not part:
                continue

            qty, qty_type = parse_qty(qty_raw)
            if qty is None:
                continue

            if part not in wh_index:
                continue

            before = remaining.get(part, 0)
            requested = before if qty == "ALL" else int(qty)
            sent = min(before, requested)
            after = before - sent
            remaining[part] = after

            if sent <= 0:
                continue

            job1 = self._pull_value(row, self.P_JOB, self.P_COL_JOB_IDX)
            job2 = self._pull_value(row, self.P_JOB2, self.P_COL_JOB2_IDX)
            rm = self._pull_value(row, self.P_RM, self.P_COL_RM_IDX)
            loc = self._pull_value(row, self.P_LOC, self.P_COL_LOC_IDX)

            rm = "" if pd.isna(rm) else str(rm).strip()
            loc = "" if pd.isna(loc) else str(loc).strip()

            job_text = self._combine_job_text(job1, job2)
            loc_full = " ".join([x for x in [rm, loc] if x and x.lower() != "nan"]).strip()

            tag_rows.append({
                "part": self._zpl_safe(part),
                "qty": sent,
                "job": self._zpl_safe(job_text),
                "loc": self._zpl_safe(loc_full),
                "date": run_date
            })

        return tag_rows

    def make_zebra_tag_zpl(self, part, qty, job, loc, date_text):
        part = self._zpl_safe(part)
        qty = self._zpl_safe(qty)
        job = self._zpl_safe(job)
        loc = self._zpl_safe(loc)
        date_text = self._zpl_safe(date_text)

        zpl = f"""
^XA
^PW812
^LL1218
^LH0,0
^CI28

^FO40,30^A0N,50,50^FDPART:^FS
^FO220,25^A0N,70,70^FD{part}^FS

^FO40,120^A0N,42,42^FDQTY:^FS
^FO170,120^A0N,42,42^FD{qty}^FS

^FO40,185^A0N,36,36^FDJOB:^FS
^FO170,185^A0N,36,36^FD{job}^FS

^FO40,245^A0N,36,36^FDLOC:^FS
^FO170,245^A0N,36,36^FD{loc}^FS

^FO40,305^A0N,30,30^FDDATE: {date_text}^FS

^FO40,390^BY3,3,120
^BCN,120,Y,N,N
^FD{part}^FS

^XZ
"""
        return zpl.strip() + "\n"

    def _build_all_zpl(self):
        tag_rows = self.build_tag_rows()
        if not tag_rows:
            return "", []

        zpl_chunks = []
        for row in tag_rows:
            zpl_chunks.append(
                self.make_zebra_tag_zpl(
                    part=row["part"],
                    qty=row["qty"],
                    job=row["job"],
                    loc=row["loc"],
                    date_text=row["date"]
                )
            )
        return "".join(zpl_chunks), tag_rows

    def export_zebra_tags(self):
        if self.warehouse_df is None or self.pull_df is None:
            messagebox.showinfo("Missing Data", "Load warehouse + pull list first.")
            return

        zpl_text, tag_rows = self._build_all_zpl()
        if not zpl_text:
            messagebox.showinfo("No Tags", "No printable tag rows were generated.\n\nThis usually means nothing was actually sent.")
            return

        path = filedialog.asksaveasfilename(
            title="Save Zebra Tag File",
            initialfile=safe_default_filename("Zebra_Tags", "zpl"),
            defaultextension=".zpl",
            filetypes=[("Zebra ZPL files", "*.zpl"), ("Text files", "*.txt"), ("All files", "*.*")]
        )
        if not path:
            return

        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write(zpl_text)
            self.status_var.set(f"Exported Zebra tags: {os.path.basename(path)}")
            messagebox.showinfo("Exported", f"Zebra tag file saved:\n{path}")
        except Exception as e:
            messagebox.showerror("Export Error", f"Could not export Zebra tags.\n\n{e}")

    def print_zebra_tags(self):
        """
        Windows:
            Uses win32print RAW output.
        Mac/Linux:
            Writes temp .zpl file and sends with lp.
        """
        if self.warehouse_df is None or self.pull_df is None:
            messagebox.showinfo("Missing Data", "Load warehouse + pull list first.")
            return

        zpl_text, tag_rows = self._build_all_zpl()
        if not zpl_text:
            messagebox.showinfo("No Tags", "No printable tag rows were generated.\n\nThis usually means nothing was actually sent.")
            return

        printer_name = self.ZEBRA_PRINTER_NAME.strip()
        if not printer_name:
            messagebox.showinfo("Printer Not Set", "Set the Zebra printer name in Settings first.")
            return

        system_name = platform.system().lower()

        if "windows" in system_name:
            try:
                import win32print

                hprinter = win32print.OpenPrinter(printer_name)
                try:
                    win32print.StartDocPrinter(hprinter, 1, ("Tekmor Zebra Tags", None, "RAW"))
                    win32print.StartPagePrinter(hprinter)
                    win32print.WritePrinter(hprinter, zpl_text.encode("utf-8"))
                    win32print.EndPagePrinter(hprinter)
                    win32print.EndDocPrinter(hprinter)
                finally:
                    win32print.ClosePrinter(hprinter)

                self.status_var.set(f"Printed {len(tag_rows)} Zebra tag(s) to {printer_name}")
                messagebox.showinfo("Printed", f"Sent {len(tag_rows)} tag(s) to:\n{printer_name}")
                return

            except ImportError:
                messagebox.showerror(
                    "Printing Not Available",
                    "Direct Windows printing needs pywin32 installed.\n\n"
                    "Run:\n"
                    "pip install pywin32\n\n"
                    "You can still use File → Export Zebra Tags (.zpl)... right now."
                )
                return
            except Exception as e:
                messagebox.showerror(
                    "Print Error",
                    f"Could not send tags to printer '{printer_name}'.\n\n{e}\n\n"
                    "You can still use File → Export Zebra Tags (.zpl)..."
                )
                return

        else:
            tmp_path = None
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".zpl", mode="w", encoding="utf-8") as tmp:
                    tmp.write(zpl_text)
                    tmp_path = tmp.name

                cmd = ["lp", "-d", printer_name, "-o", "raw", tmp_path]
                result = subprocess.run(cmd, capture_output=True, text=True)

                if result.returncode != 0:
                    raise RuntimeError(result.stderr.strip() or result.stdout.strip() or "Unknown lp error")

                self.status_var.set(f"Printed {len(tag_rows)} Zebra tag(s) to {printer_name}")
                messagebox.showinfo("Printed", f"Sent {len(tag_rows)} tag(s) to:\n{printer_name}")

            except Exception as e:
                messagebox.showerror(
                    "Print Error",
                    f"Could not send tags to printer '{printer_name}'.\n\n{e}\n\n"
                    "Make sure the printer name matches exactly what macOS shows in Printers & Scanners.\n"
                    "You can still use File → Export Zebra Tags (.zpl)..."
                )
            finally:
                if tmp_path and os.path.exists(tmp_path):
                    try:
                        os.remove(tmp_path)
                    except Exception:
                        pass

    # ---------- Outputs ----------
    def _show_shipment_summary(self, shipment_rows):
        if shipment_rows:
            ship_df = pd.DataFrame(shipment_rows)
            lines = []
            lines.append(f"Shipment Summary — Date: {today_mdy2()}")
            lines.append("-" * 60)
            lines.append(ship_df.to_string(index=False))
            self._set_shipment("\n".join(lines))
        else:
            self._set_shipment("Nothing was sent (all parts were 0 on hand or skipped/not found).")

    def _post_run_outputs(self, plan_df, pre_log_df, shortages_df):
        batch_log = pre_log_df.copy()
        self.preview_pull_list(summary=self._summarize_batch_log(batch_log))

        shipment_rows = []
        if plan_df is not None and not plan_df.empty:
            for _, r in plan_df.iterrows():
                shipment_rows.append({
                    "part": r["part"],
                    "sent": int(r["sent"]),
                    "before": int(r["before"]),
                    "after": int(r["after"])
                })
        self._show_shipment_summary(shipment_rows)
        self._last_shortages_df = shortages_df

    def _summarize_batch_log(self, log_df: pd.DataFrame) -> dict:
        if log_df is None or log_df.empty:
            return {
                "changed": 0,
                "not_found": 0,
                "partial_fill": 0,
                "skipped": 0,
                "manual_receive": 0,
                "manual_send": 0,
            }

        counts = log_df["category"].value_counts().to_dict()
        return {
            "changed": int(counts.get("changed", 0)),
            "not_found": int(counts.get("not_found", 0)),
            "partial_fill": int(counts.get("partial_fill", 0)),
            "skipped": int(counts.get("skipped", 0)),
            "manual_receive": int(counts.get("manual_receive", 0)),
            "manual_send": int(counts.get("manual_send", 0)),
        }

    def save_updated_warehouse(self):
        if self.warehouse_df is None:
            messagebox.showinfo("No Data", "Nothing to save. Load a warehouse file first.")
            return
        path = filedialog.asksaveasfilename(
            title="Save Updated Warehouse As",
            initialfile=safe_default_filename("Warehouse_UPDATED"),
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")]
        )
        if not path:
            return
        try:
            self.warehouse_df.to_csv(path, index=False)
            self.status_var.set(f"Saved updated warehouse: {os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("Save Error", f"Could not save file.\n\n{e}")

    def export_log(self):
        if self.log_df is None or self.log_df.empty:
            messagebox.showinfo("No Log", "No log data yet.")
            return
        path = filedialog.asksaveasfilename(
            title="Export Log As",
            initialfile=safe_default_filename("Inventory_Log"),
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")]
        )
        if not path:
            return
        try:
            self.log_df.to_csv(path, index=False)
            self.status_var.set(f"Exported log: {os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("Export Error", f"Could not export log.\n\n{e}")

    def export_shortages(self):
        if self._last_shortages_df is None or self._last_shortages_df.empty:
            messagebox.showinfo("No Shortages", "No shortages to export yet.\nRun a pull (or Dry Run) first.")
            return

        path = filedialog.asksaveasfilename(
            title="Export Shortages As",
            initialfile=safe_default_filename("Shortages"),
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")]
        )
        if not path:
            return
        try:
            self._last_shortages_df.to_csv(path, index=False)
            self.status_var.set(f"Exported shortages: {os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("Export Error", f"Could not export shortages.\n\n{e}")

    # ---------- Search ----------
    def search_part(self):
        if self.warehouse_df is None:
            messagebox.showinfo("No Data", "Load a warehouse file first.")
            return

        query = normalize_part(self.search_var.get())
        if not query:
            return

        if self.W_PART not in self.warehouse_df.columns:
            messagebox.showerror("Config Error", "Warehouse part column not set correctly (Tools > Settings).")
            return

        matches = self.warehouse_df[self.warehouse_df[self.W_PART] == query]
        if matches.empty:
            self._set_details(f"❌ Part not found: {query}")
            return

        row = matches.iloc[0].to_dict()
        lines = [f"✅ Part Found: {query}", "-" * 40]
        for k, v in row.items():
            lines.append(f"{k}: {v}")
        self._set_details("\n".join(lines))
        self._select_part_in_tree(query)

    def clear_search(self):
        self.search_var.set("")
        self._set_details("")

    # ---------- Inventory filters ----------
    def filter_low_stock(self):
        if self.warehouse_df is None:
            return
        if self.W_QTY not in self.warehouse_df.columns:
            return
        df = self.warehouse_df[self.warehouse_df[self.W_QTY] <= 0]
        self.refresh_inventory_view(df_to_show=df)
        self.status_var.set("Showing low stock (<= 0).")

    def clear_filters(self):
        self.refresh_inventory_view()
        self.status_var.set("Filters cleared.")

    # ---------- Pull preview ----------
    def preview_pull_list(self, summary=None):
        if self.pull_df is None:
            self._set_pull_preview("No pull list loaded.")
            return

        lines = []
        if summary:
            lines.append("Last Apply Summary")
            lines.append("-" * 28)
            for k, v in summary.items():
                lines.append(f"{k}: {v}")
            lines.append("\nPull List (first 15 rows)")
            lines.append("-" * 28)
        else:
            lines.append("Pull List (first 15 rows)")
            lines.append("-" * 28)

        try:
            lines.append(self.pull_df.head(15).to_string(index=False))
        except Exception:
            lines.append(str(self.pull_df.head(15)))

        self._set_pull_preview("\n".join(lines))

    # ---------- Settings ----------
    def open_settings(self):
        win = tk.Toplevel(self)
        win.title("Settings")
        win.geometry("650x650")
        win.grab_set()

        frm = ttk.Frame(win, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frm, text="Column Mapping", font=("Segoe UI", 11, "bold")).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 10)
        )

        ttk.Label(frm, text="Warehouse Part Column:").grid(row=1, column=0, sticky="w")
        w_part = tk.StringVar(value=self.W_PART)
        ttk.Entry(frm, textvariable=w_part).grid(row=1, column=1, sticky="ew")

        ttk.Label(frm, text="Warehouse Qty Column:").grid(row=2, column=0, sticky="w")
        w_qty = tk.StringVar(value=self.W_QTY)
        ttk.Entry(frm, textvariable=w_qty).grid(row=2, column=1, sticky="ew")

        ttk.Label(frm, text="Warehouse MAIN Date Column (Date):").grid(row=3, column=0, sticky="w")
        w_date_main = tk.StringVar(value=self.W_DATE_MAIN)
        ttk.Entry(frm, textvariable=w_date_main).grid(row=3, column=1, sticky="ew")

        ttk.Label(frm, text="Warehouse Last Updated Column:").grid(row=4, column=0, sticky="w")
        w_date = tk.StringVar(value=self.W_DATE)
        ttk.Entry(frm, textvariable=w_date).grid(row=4, column=1, sticky="ew")

        ttk.Label(frm, text="Warehouse In/Out Column:").grid(row=5, column=0, sticky="w")
        w_inout = tk.StringVar(value=self.W_INOUT)
        ttk.Entry(frm, textvariable=w_inout).grid(row=5, column=1, sticky="ew")

        ttk.Separator(frm).grid(row=6, column=0, columnspan=2, sticky="ew", pady=12)

        ttk.Label(frm, text="Pull Part Column:").grid(row=7, column=0, sticky="w")
        p_part = tk.StringVar(value=self.P_PART)
        ttk.Entry(frm, textvariable=p_part).grid(row=7, column=1, sticky="ew")

        ttk.Label(frm, text="Pull Qty Column:").grid(row=8, column=0, sticky="w")
        p_qty = tk.StringVar(value=self.P_QTY)
        ttk.Entry(frm, textvariable=p_qty).grid(row=8, column=1, sticky="ew")

        ttk.Separator(frm).grid(row=9, column=0, columnspan=2, sticky="ew", pady=12)

        ttk.Label(frm, text="Pull Job 1 Column:").grid(row=10, column=0, sticky="w")
        p_job = tk.StringVar(value=self.P_JOB)
        ttk.Entry(frm, textvariable=p_job).grid(row=10, column=1, sticky="ew")

        ttk.Label(frm, text="Pull Job 2 Column:").grid(row=11, column=0, sticky="w")
        p_job2 = tk.StringVar(value=self.P_JOB2)
        ttk.Entry(frm, textvariable=p_job2).grid(row=11, column=1, sticky="ew")

        ttk.Label(frm, text="Pull RM Column:").grid(row=12, column=0, sticky="w")
        p_rm = tk.StringVar(value=self.P_RM)
        ttk.Entry(frm, textvariable=p_rm).grid(row=12, column=1, sticky="ew")

        ttk.Label(frm, text="Pull Location Column:").grid(row=13, column=0, sticky="w")
        p_loc = tk.StringVar(value=self.P_LOC)
        ttk.Entry(frm, textvariable=p_loc).grid(row=13, column=1, sticky="ew")

        ttk.Label(frm, text="Zebra Printer Name:").grid(row=14, column=0, sticky="w")
        zebra_name = tk.StringVar(value=self.ZEBRA_PRINTER_NAME)
        ttk.Entry(frm, textvariable=zebra_name).grid(row=14, column=1, sticky="ew")

        frm.columnconfigure(1, weight=1)

        def save_settings():
            self.W_PART = w_part.get().strip()
            self.W_QTY = w_qty.get().strip()
            self.W_DATE_MAIN = w_date_main.get().strip()
            self.W_DATE = w_date.get().strip()
            self.W_INOUT = w_inout.get().strip()

            self.P_PART = p_part.get().strip()
            self.P_QTY = p_qty.get().strip()
            self.P_JOB = p_job.get().strip()
            self.P_JOB2 = p_job2.get().strip()
            self.P_RM = p_rm.get().strip()
            self.P_LOC = p_loc.get().strip()
            self.ZEBRA_PRINTER_NAME = zebra_name.get().strip()

            messagebox.showinfo("Saved", "Settings updated.")
            win.destroy()

        ttk.Button(frm, text="Save", command=save_settings).grid(row=15, column=0, pady=14, sticky="w")
        ttk.Button(frm, text="Cancel", command=win.destroy).grid(row=15, column=1, pady=14, sticky="e")

    def about(self):
        messagebox.showinfo(
            "About",
            "Tekmor Warehouse Tool\n\n"
            "Includes:\n"
            "- Safe pull processing\n"
            "- No negative inventory\n"
            "- Date only updates when qty actually moves\n"
            "- Manual Receive (+)\n"
            "- Manual Send (-)\n"
            "- Dry Run / backups / logs / shortages\n"
            "- Paste full pull list directly into app\n"
            "- Combines Job 1 + Job 2 on Zebra tags\n"
            "- Zebra ZT230 tag export / print (Windows + Mac/Linux)\n"
        )

    # ---------- Tree utils ----------
    def _clear_tree(self):
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = []

    def _populate_tree(self, df: pd.DataFrame):
        self._clear_tree()

        cols = list(df.columns)
        self.tree["columns"] = cols

        for col in cols:
            self.tree.heading(col, text=col)
            width = 120
            if len(col) > 14:
                width = 170
            self.tree.column(col, width=width, stretch=True)

        max_rows = min(len(df), 5000)
        for i in range(max_rows):
            values = [df.iloc[i][c] for c in cols]
            self.tree.insert("", "end", values=values)

        if len(df) > max_rows:
            self.status_var.set(f"Showing first {max_rows} rows (of {len(df)}).")

    def _select_part_in_tree(self, part_value: str):
        cols = self.tree["columns"]
        if self.W_PART not in cols:
            return
        part_col_index = cols.index(self.W_PART)

        for item in self.tree.get_children():
            vals = self.tree.item(item, "values")
            if len(vals) > part_col_index and normalize_part(vals[part_col_index]) == part_value:
                self.tree.selection_set(item)
                self.tree.see(item)
                break

    def on_row_select(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        item = sel[0]
        vals = self.tree.item(item, "values")
        cols = self.tree["columns"]
        row_dict = {cols[i]: vals[i] for i in range(min(len(cols), len(vals)))}

        lines = ["Selected Row", "-" * 40]
        for k, v in row_dict.items():
            lines.append(f"{k}: {v}")
        self._set_details("\n".join(lines))


if __name__ == "__main__":
    app = WarehouseApp()
    app.mainloop()
