# main.py
import os
import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import pandas as pd
from openpyxl.workbook.protection import WorkbookProtection

ALLOWED_INPUT = [("Data files", "*.csv *.tsv *.xlsx *.xls *.json"), ("All files", "*.*")]
ALLOWED_DOSE  = [("Dose dict", "*.csv *.tsv *.xlsx *.xls *.json"), ("All files", "*.*")]

# Lunghezza standard per SUBJECT numerici (padding con zeri a sinistra)
SUBJECT_PAD_TO_LEN = 8  # cambia se necessario; metti None per disattivare il padding

# ---------- helpers ----------
def friendly_path(p):
    home = os.path.expanduser("~")
    return p.replace(home, "~") if p else ""

def normalize_header(name: str) -> str:
    s = str(name).replace("\u00A0", " ")
    s = s.strip()
    s = re.sub(r"[\s\u00A0]+", " ", s)
    return s.lower()

def nospace_key(s: str) -> str:
    return s.replace(" ", "")

def _clean_cell(x):
    if pd.isna(x):
        return pd.NA
    s = str(x).replace("\u00A0", " ").strip()
    return pd.NA if s == "" else s

def _keyize_id(x: object) -> str:
    # chiave normalizzata: se tutta numerica, rimuove gli zeri a sinistra (00123 -> 123)
    if pd.isna(x):
        return ""
    s = str(x).replace("\u00A0", " ").strip()
    if s.isdigit():
        s = s.lstrip("0")
        if s == "":
            s = "0"
    return s

def _pad_if_numeric(s: str, pad_len: int | None) -> str:
    if s is None:
        return ""
    s = str(s)
    if pad_len and s.isdigit() and len(s) < pad_len:
        return s.zfill(pad_len)
    return s

def _display_for_key(k: str, display_map: dict) -> str:
    """
    Restituisce la display string per la chiave k.
    Se non c'è in mappa ed è numerica, applica padding a SUBJECT_PAD_TO_LEN.
    """
    if k is None:
        return ""
    k = str(k)
    disp = display_map.get(k)
    if disp is not None and disp != "":
        return _pad_if_numeric(disp, SUBJECT_PAD_TO_LEN)
    return _pad_if_numeric(k, SUBJECT_PAD_TO_LEN)

def read_table(path: str) -> pd.DataFrame:
    p = path.lower()
    try:
        if p.endswith(".csv"):
            return pd.read_csv(path, dtype=str)
        if p.endswith((".tsv", ".tab")):
            return pd.read_csv(path, sep="\t", dtype=str)
        if p.endswith((".xlsx", ".xls")):
            return pd.read_excel(path, dtype=str)
        if p.endswith(".json"):
            df = pd.read_json(path, orient="records")
            return df.applymap(lambda v: None if pd.isna(v) else str(v))
    except Exception:
        pass
    try:
        return pd.read_csv(path, dtype=str)
    except Exception:
        return pd.read_excel(path, dtype=str)

def match_columns(df: pd.DataFrame, required_keys: dict) -> dict:
    norm_map = {}
    for col in df.columns:
        norm = nospace_key(normalize_header(col))
        norm_map.setdefault(norm, col)

    resolved, missing = {}, []
    for logical, candidates in required_keys.items():
        found = None
        for cand in candidates:
            key = nospace_key(normalize_header(cand))
            if key in norm_map:
                found = norm_map[key]; break
        if found is None:
            missing.append(logical)
        else:
            resolved[logical] = found

    if missing:
        seen = ", ".join([f"'{c}'" for c in df.columns])
        need = ", ".join(missing)
        raise ValueError(
            f"Missing required column(s): {need}. "
            f"Available columns: {seen}. Headers are matched case-insensitively and ignore spaces/NBSP."
        )
    return resolved

def find_columns_regex(df: pd.DataFrame, pattern: str) -> list[str]:
    cols = []
    for col in df.columns:
        norm = nospace_key(normalize_header(col))
        if re.fullmatch(pattern, norm):
            cols.append(col)
    return cols

ILLEGAL_SHEET_CHARS = r'[/\\*?:\[\]]'
def sanitize_sheet_name(name: str, used: set, prefix: str = "") -> str:
    base = "BLANK" if name is None or str(name).strip() == "" else str(name)
    base = re.sub(ILLEGAL_SHEET_CHARS, "_", base)
    base = re.sub(r"[\r\n]+", " ", base).strip()
    if prefix:
        base = f"{prefix}{base}"
    base = base[:31] if len(base) > 31 else base
    if base == "": base = "SHEET"
    candidate = base
    i = 2
    while candidate in used:
        suffix = f" ({i})"
        candidate = (base[: (31 - len(suffix))] + suffix) if len(base) + len(suffix) > 31 else base + suffix
        i += 1
    used.add(candidate)
    return candidate

# ---------- percent parsing ----------
def _to_percent(value):
    if pd.isna(value): return pd.NA
    s = str(value).replace("\u00A0", " ").strip()
    if s == "": return pd.NA
    s = s.replace("%", "")
    try: v = float(s)
    except Exception: return pd.NA
    if v <= 1.0: v = v * 100.0
    if v < 0: v = 0.0
    if v > 1000: v = 100.0
    return v

# ---------- canonical SUBJECT display ----------
def _norm_disp(x: object) -> str:
    return "" if pd.isna(x) else str(x).replace("\u00A0", " ").strip()

def _pick_display(disp_from_dose: str, disp_from_input: str) -> str:
    a = _norm_disp(disp_from_dose)
    b = _norm_disp(disp_from_input)
    if a.isdigit() and b.isdigit():
        return a if len(a) >= len(b) else b  # più zeri preservati
    if a.isdigit() and not b.isdigit():
        return b or a
    if b.isdigit() and not a.isdigit():
        return a or b
    return b or a

def build_canonical_display_map(dose_df: pd.DataFrame, input_df: pd.DataFrame, col_patient: str, col_subject: str) -> dict:
    dose_tmp = pd.DataFrame({
        "key": dose_df[col_patient].map(_keyize_id),
        "disp": dose_df[col_patient].map(_clean_cell),
    }).dropna(subset=["key"]).drop_duplicates("key")

    input_tmp = pd.DataFrame({
        "key": input_df[col_subject].map(_keyize_id),
        "disp": input_df[col_subject].map(_clean_cell),
    }).dropna(subset=["key"]).drop_duplicates("key")

    merged = dose_tmp.merge(input_tmp, on="key", how="outer", suffixes=("_dose","_input"))
    canon = []
    for a, b in zip(merged.get("disp_dose"), merged.get("disp_input")):
        chosen = _pick_display(a, b)
        chosen = _pad_if_numeric(chosen, SUBJECT_PAD_TO_LEN)  # padding se serve
        canon.append(chosen)
    merged["canon"] = canon
    merged["canon"] = merged["canon"].fillna(merged.get("disp_dose")).fillna(merged.get("disp_input")).fillna("")
    merged["canon"] = merged["canon"].map(lambda s: _pad_if_numeric("" if pd.isna(s) else str(s), SUBJECT_PAD_TO_LEN))
    return {row["key"]: row["canon"] for _, row in merged.iterrows() if row["key"] != ""}

def extract_input_keys(input_df: pd.DataFrame, subj_col: str) -> set[str]:
    """Restituisce l'insieme delle chiavi normalizzate presenti nell'input."""
    return set(input_df[subj_col].map(_keyize_id).dropna().tolist())

# ---------- core processing ----------
def prepare_groups(dose_df: pd.DataFrame, col_patient: str, col_group: str, valid_keys: set[str]) -> dict[str, list[str]]:
    """
    Ritorna: group -> lista unica ordinata di KEY normalizzate,
    **solo** per le chiavi presenti nell'INPUT (valid_keys).
    """
    keys   = dose_df[col_patient].map(_keyize_id)
    groups = dose_df[col_group].map(_clean_cell)

    tmp = pd.DataFrame({"_KEY": keys, "__group__": groups})
    tmp = tmp.dropna(subset=["__group__"])  # drop gruppi vuoti
    tmp = tmp[tmp["_KEY"] != ""]
    tmp = tmp[tmp["_KEY"].isin(valid_keys)]  # filtro: solo soggetti presenti nell'input

    result: dict[str, list[str]] = {}
    for grp, subdf in tmp.groupby("__group__", dropna=False):
        ordered, seen = [], set()
        for k in subdf["_KEY"]:
            if k not in seen:
                seen.add(k)
                ordered.append(k)
        if ordered:
            result[grp] = ordered
    return result

def build_group_tables(input_df: pd.DataFrame, input_cols: dict, groups: dict, display_map: dict) -> dict:
    """
    Ritorna {group -> DataFrame con [SUBJECT, VISIT, BMBLASTS, CD123, CD123+ BLASTS (%)]},
    usando SEMPRE display_map per SUBJECT. **Niente filler**: solo righe presenti nell'input.
    """
    subj_col  = input_cols["subject"]
    visit_col = input_cols["visit"]
    bm_col    = input_cols["bmblasts"]
    cd_col    = input_cols["cd123"]

    df_in = input_df.copy()
    df_in["_SUBJECT_KEY"] = df_in[subj_col].map(_keyize_id)

    df_in_out = df_in[["_SUBJECT_KEY", visit_col, bm_col, cd_col]].copy()
    df_in_out.rename(columns={visit_col:"VISIT", bm_col:"BMBLASTS", cd_col:"CD123"}, inplace=True)

    bm_pct = df_in_out["BMBLASTS"].map(_to_percent)
    cd_pct = df_in_out["CD123"].map(_to_percent)
    with pd.option_context('mode.use_inf_as_na', True):
        df_in_out["CD123+ BLASTS (%)"] = (bm_pct * (cd_pct / 100.0)).round(1)

    out = {}
    for grp, keys in groups.items():
        key_set = set(keys)
        part = df_in_out.loc[df_in_out["_SUBJECT_KEY"].isin(key_set),
                             ["_SUBJECT_KEY","VISIT","BMBLASTS","CD123","CD123+ BLASTS (%)"]].copy()

        # SUBJECT dalla mappa canonica (con eventuale padding)
        part["SUBJECT"] = part["_SUBJECT_KEY"].map(lambda k: _display_for_key("" if pd.isna(k) else str(k), display_map))

        # Ordina
        if "VISIT" in part.columns:
            part.sort_values(by=["SUBJECT","VISIT"], inplace=True, kind="mergesort")
        else:
            part.sort_values(by=["SUBJECT"], inplace=True, kind="mergesort")

        part = part[["SUBJECT","VISIT","BMBLASTS","CD123","CD123+ BLASTS (%)"]].reset_index(drop=True)
        out[grp] = part
    return out

def is_tp53_positive(value) -> bool:
    if pd.isna(value): return False
    s = str(value).replace("\u00A0"," ").strip().lower()
    return s == "positive"

def compute_tp53_summary(groups: dict, input_df: pd.DataFrame, subj_col_actual: str, tp53_col_actual: str) -> pd.DataFrame:
    tmp = input_df[[subj_col_actual, tp53_col_actual]].copy()
    tmp["_KEY"] = tmp[subj_col_actual].map(_keyize_id)
    tmp["__mut__"] = tmp[tp53_col_actual].map(is_tp53_positive)
    mut_map = tmp.groupby("_KEY")["__mut__"].any().to_dict()

    rows = []
    for grp, keys in groups.items():
        n_pat = len(keys)
        mut_count = sum(1 for k in keys if mut_map.get(k, False))
        perc = round((mut_count / n_pat * 100.0), 2) if n_pat > 0 else 0.0
        rows.append({"Cohort": grp, "N patients": n_pat, "TP53 mutated": mut_count, "Percentage": perc})
    return pd.DataFrame(rows).sort_values(by="Cohort", kind="mergesort").reset_index(drop=True)

# ---------- TP53: tutte le coppie DNA/PROTEIN (N=1..∞), senza duplicati ----------
def extract_tp53_labels_map(input_df: pd.DataFrame, subj_col: str):
    df = input_df.copy()
    df["_SUBJECT_KEY"] = df[subj_col].map(_keyize_id)

    norm_to_actual = {nospace_key(normalize_header(c)): c for c in df.columns}
    idx_set = set()
    num_re = re.compile(r"tp53(?:dna|protein)(\d+)")
    for c in df.columns:
        m = num_re.fullmatch(nospace_key(normalize_header(c)))
        if m: idx_set.add(int(m.group(1)))
    idx_list = sorted(idx_set)

    def _pair_label_simple(dna, prot):
        dna_s  = "" if pd.isna(dna) else str(dna).replace("\u00A0"," ").strip()
        prot_s = "" if pd.isna(prot) else str(prot).replace("\u00A0"," ").strip()
        if dna_s == "" and prot_s == "": return None
        if dna_s == "" or prot_s == "":  return f"{dna_s} | {prot_s}".strip()
        return f"{dna_s} | {prot_s}"

    subj_labels = {}  # key -> set(labels)
    for _, row in df.iterrows():
        key = row["_SUBJECT_KEY"]
        if key == "": continue
        labels = subj_labels.setdefault(key, set())
        for n in idx_list:
            dna_col  = norm_to_actual.get(f"tp53dna{n}")
            prot_col = norm_to_actual.get(f"tp53protein{n}")
            dna_v  = row.get(dna_col) if dna_col in df.columns else pd.NA
            prot_v = row.get(prot_col) if prot_col in df.columns else pd.NA
            lab = _pair_label_simple(dna_v, prot_v)
            if lab: labels.add(lab)
    return subj_labels

def build_tp53_matrices_per_group(groups: dict, subj_labels: dict, display_map: dict) -> dict:
    out = {}
    for grp, keys in groups.items():
        # tutte le etichette presenti nella coorte
        label_set = set()
        for k in keys:
            label_set |= subj_labels.get(k, set())
        labels_sorted = sorted(label_set)

        if labels_sorted:
            data = []
            for k in keys:
                labs = subj_labels.get(k, set())
                row_marks = {lab: ("X" if lab in labs else "") for lab in labels_sorted}
                row_marks["SUBJECT"] = _display_for_key(k, display_map)
                data.append(row_marks)
            cols = ["SUBJECT"] + labels_sorted
            mat = pd.DataFrame(data, columns=cols)
        else:
            mat = pd.DataFrame({"SUBJECT": [_display_for_key(k, display_map) for k in keys]})
        out[grp] = mat
    return out

# ---------- writer (SUBJECT testo + protezione opzionale) ----------
def write_excel_all(group_tables: dict, tp53_summary: pd.DataFrame, tp53_mats: dict, out_path: str, password: str | None):
    used = set()
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        # Per-coorte
        written_group_sheetnames = []
        for grp, df_pat in group_tables.items():
            sheet = sanitize_sheet_name(grp, used)
            df_pat.to_excel(writer, sheet_name=sheet, index=False)
            written_group_sheetnames.append(sheet)

        # TP53 summary globale
        tp53_summary.to_excel(writer, sheet_name="TP53 mutation", index=False)

        # Matrici TP53 per coorte
        written_tp53_mat_sheetnames = []
        for grp, mat in tp53_mats.items():
            sheet = sanitize_sheet_name(grp, used, prefix="TP53 muts - ")
            mat.to_excel(writer, sheet_name=sheet, index=False)
            written_tp53_mat_sheetnames.append(sheet)

        # post-processing con openpyxl
        wb = writer.book

        # SUBJECT come testo (colonna A) su fogli gruppi e TP53-muts
        def force_subject_text(ws):
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=1):
                cell = row[0]
                cell.number_format = "@"
                if cell.value is None:
                    continue
                cell.value = "" if pd.isna(cell.value) else str(cell.value)

        for name in written_group_sheetnames + written_tp53_mat_sheetnames:
            ws = wb[name]
            force_subject_text(ws)

        # Protezione opzionale
        if password and password.strip() != "":
            try:
                wb.security = WorkbookProtection(lockStructure=True)
                try:
                    wb.security.set_workbook_password(password)
                except Exception:
                    wb.security.workbookPassword = password
            except Exception:
                pass
            for ws in wb.worksheets:
                ws.protection.sheet = True
                try:
                    ws.protection.set_password(password)
                except Exception:
                    ws.protection.password = password

# ---------- GUI ----------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        try:
            os.chdir(os.path.expanduser("~"))
        except Exception:
            pass

        self.title("FDA_Tabulation")
        self.geometry("900x420")
        self.resizable(False, False)

        self.last_dir = os.path.expanduser("~")
        self.input_path = tk.StringVar()
        self.dose_path  = tk.StringVar()
        self.out_dir    = tk.StringVar()
        self.out_name   = tk.StringVar(value="dose_groups.xlsx")

        pad = {"padx": 12, "pady": 8}
        frm = ttk.Frame(self)
        frm.pack(fill="both", expand=True, **pad)
        frm.columnconfigure(1, weight=1)

        ttk.Label(frm, text="Input data file (with SUBJECT):").grid(row=0, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.input_path, width=70).grid(row=0, column=1, sticky="we", padx=(8,8))
        ttk.Button(frm, text="Browse…", command=self.browse_input).grid(row=0, column=2, sticky="e")

        ttk.Label(frm, text="Patient dose dictionary:").grid(row=1, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.dose_path, width=70).grid(row=1, column=1, sticky="we", padx=(8,8))
        ttk.Button(frm, text="Browse…", command=self.browse_dose).grid(row=1, column=2, sticky="e")

        ttk.Label(frm, text="Output folder:").grid(row=2, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.out_dir, width=70).grid(row=2, column=1, sticky="we", padx=(8,8))
        ttk.Button(frm, text="Choose…", command=self.browse_outdir).grid(row=2, column=2, sticky="e")

        ttk.Label(frm, text="Output file name (.xlsx):").grid(row=3, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.out_name, width=40).grid(row=3, column=1, sticky="w", padx=(8,8))

        btns = ttk.Frame(frm)
        btns.grid(row=5, column=0, columnspan=3, sticky="e", pady=(16,0))
        ttk.Button(btns, text="Quit", command=self.destroy).pack(side="right", padx=(8,0))
        ttk.Button(btns, text="Run", command=self.run).pack(side="right")

    def browse_input(self):
        path = filedialog.askopenfilename(title="Select Input Data", initialdir=self.last_dir, filetypes=ALLOWED_INPUT)
        if path:
            self.last_dir = os.path.dirname(path)
            self.input_path.set(path)

    def browse_dose(self):
        path = filedialog.askopenfilename(title="Select Patient Dose Dictionary", initialdir=self.last_dir, filetypes=ALLOWED_DOSE)
        if path:
            self.last_dir = os.path.dirname(path)
            self.dose_path.set(path)

    def browse_outdir(self):
        path = filedialog.askdirectory(title="Select Output Folder", initialdir=self.last_dir)
        if path:
            self.last_dir = path
            self.out_dir.set(path)

    def run(self):
        in_path   = self.input_path.get().strip()
        dose_path = self.dose_path.get().strip()
        out_dir   = self.out_dir.get().strip()
        out_name  = self.out_name.get().strip()

        if not in_path or not os.path.isfile(in_path):
            messagebox.showerror("Missing input", "Please select a valid INPUT data file (with SUBJECT)."); return
        if not dose_path or not os.path.isfile(dose_path):
            messagebox.showerror("Missing dose dict", "Please select a valid patient dose dictionary file."); return
        if not out_dir or not os.path.isdir(out_dir):
            messagebox.showerror("Missing output folder", "Please choose a valid output folder."); return
        if out_name == "":
            messagebox.showerror("Missing output name", "Please provide an output file name (e.g., dose_groups.xlsx)."); return
        if not out_name.lower().endswith(".xlsx"):
            out_name += ".xlsx"
        out_path = os.path.join(out_dir, out_name)

        # Protezione opzionale
        protect = messagebox.askyesno("Protection", "Protect workbook structure and all sheets with a password?")
        password = None
        if protect:
            password = simpledialog.askstring("Password", "Enter password:", show="*")
            if password is None:
                messagebox.showinfo("Cancelled", "Export cancelled."); return

        try:
            dose_df  = read_table(dose_path)
            input_df = read_table(in_path)

            dose_cols = match_columns(dose_df, {"patientnumber":["Patient Number"], "treatmentgroup":["Treatment Group"]})
            col_patient = dose_cols["patientnumber"]; col_group = dose_cols["treatmentgroup"]

            input_cols = match_columns(
                input_df,
                {
                    "subject":["SUBJECT"],
                    "visit":["VISIT"],
                    "bmblasts":["BMBLASTS","BM BLASTS"],
                    "cd123":["CD123"],
                    "tp53":["TP53FINALRESULT","TP53 FINAL RESULT"],
                },
            )

            # 1) chiavi presenti nell'input
            input_keys = extract_input_keys(input_df, input_cols["subject"])

            # 2) gruppi filtrati: solo pazienti presenti nell'input (e con gruppo non vuoto)
            groups = prepare_groups(dose_df, col_patient, col_group, valid_keys=input_keys)

            if not groups:
                raise ValueError("No groups found after filtering to input subjects. Check the dose dictionary and input alignment.")

            # 3) mappa display canonica (preferisce input; padding numerici a SUBJECT_PAD_TO_LEN)
            display_map = build_canonical_display_map(dose_df, input_df, col_patient, input_cols["subject"])

            # 4) tabelle per coorte (niente filler)
            group_tables = build_group_tables(input_df, input_cols, groups, display_map)

            # 5) TP53 summary (solo soggetti dei gruppi filtrati)
            tp53_df = compute_tp53_summary(groups, input_df, input_cols["subject"], input_cols["tp53"])

            # 6) Matrici TP53 per coorte (qualsiasi numero di coppie DNA/PROTEIN)
            subj_labels = extract_tp53_labels_map(input_df, input_cols["subject"])
            tp53_mats = build_tp53_matrices_per_group(groups, subj_labels, display_map)

            # 7) Scrittura Excel + SUBJECT testo + protezione
            write_excel_all(group_tables, tp53_df, tp53_mats, out_path, password=password)

        except Exception as e:
            messagebox.showerror("Processing error", str(e)); return

        messagebox.showinfo("Done", f"Excel written to:\n{friendly_path(out_path)}")

# ---------- entry ----------
if __name__ == "__main__":
    app = App()
    app.mainloop()
