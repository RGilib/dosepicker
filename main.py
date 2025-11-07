# main.py
import os
import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd

ALLOWED_INPUT = [("Data files", "*.csv *.tsv *.xlsx *.xls *.json"), ("All files", "*.*")]
ALLOWED_DOSE  = [("Dose dict", "*.csv *.tsv *.xlsx *.xls *.json"), ("All files", "*.*")]

# ---------- helpers ----------
def friendly_path(p):
    home = os.path.expanduser("~")
    return p.replace(home, "~") if p else ""

def normalize_header(name: str) -> str:
    s = str(name)
    s = s.strip()
    s = re.sub(r"\s+", " ", s)
    return s.lower()

def nospace_key(s: str) -> str:
    return s.replace(" ", "")

def _clean_cell(x):
    """Trim strings; keep NaN as NaN; convert empty-after-trim to NaN."""
    if pd.isna(x):
        return pd.NA
    s = str(x).strip()
    return pd.NA if s == "" else s

def read_table(path: str) -> pd.DataFrame:
    p = path.lower()
    if p.endswith(".csv"):  return pd.read_csv(path)
    if p.endswith((".tsv", ".tab")): return pd.read_csv(path, sep="\t")
    if p.endswith((".xlsx", ".xls")): return pd.read_excel(path)
    if p.endswith(".json"):
        df = pd.read_json(path, orient="records")
        return df if isinstance(df, pd.DataFrame) else pd.DataFrame(df)
    try:
        return pd.read_csv(path)
    except Exception:
        return pd.read_excel(path)

def match_columns(df: pd.DataFrame, required_keys: dict) -> dict:
    """
    required_keys: {"logicalname": ["candidate1", ...]}
    Returns map logicalname -> actual df column (required; raises if missing).
    """
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
                found = norm_map[key]
                break
        if found is None:
            missing.append(logical)
        else:
            resolved[logical] = found

    if missing:
        seen = ", ".join([f"'{c}'" for c in df.columns])
        need = ", ".join(missing)
        raise ValueError(
            f"Missing required column(s): {need}. "
            f"Available columns: {seen}. "
            "Headers are matched case-insensitively and ignore leading/trailing spaces."
        )
    return resolved

ILLEGAL_SHEET_CHARS = r'[/\\*?:\[\]]'
def sanitize_sheet_name(name: str, used: set) -> str:
    base = "BLANK" if name is None or str(name).strip() == "" else str(name)
    base = re.sub(ILLEGAL_SHEET_CHARS, "_", base)
    base = re.sub(r"[\r\n]+", " ", base).strip()
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

# ---------- core processing ----------
def prepare_groups(dose_df: pd.DataFrame, col_patient: str, col_group: str) -> dict:
    """
    Build {group -> ordered unique list of SUBJECTs}, dropping blank groups first.
    """
    patients = dose_df[col_patient].map(_clean_cell)
    groups   = dose_df[col_group].map(_clean_cell)

    tmp = pd.DataFrame({"SUBJECT": patients, "__group__": groups})
    tmp = tmp.dropna(subset=["__group__"])  # drop blank Treatment Group rows
    tmp = tmp.dropna(subset=["SUBJECT"])    # drop rows without subject

    result = {}
    for grp, sub in tmp.groupby("__group__", dropna=False):
        seen, ordered = set(), []
        for v in sub["SUBJECT"]:
            if v not in seen:
                seen.add(v)
                ordered.append(v)
        result[grp] = ordered
    return result

def build_group_tables(input_df: pd.DataFrame, input_cols: dict, groups: dict) -> dict:
    """
    Returns {group -> DataFrame with columns [SUBJECT, VISIT, BMBLASTS, PBBLASTS, CD123]}.
    Ensures subjects with no input rows appear with blanks.
    """
    subj_col  = input_cols["subject"]
    visit_col = input_cols["visit"]
    bm_col    = input_cols["bmblasts"]
    pb_col    = input_cols["pbblasts"]
    cd_col    = input_cols["cd123"]

    df_in = input_df.copy()
    df_in["_SUBJECT_CLEAN"] = df_in[subj_col].map(_clean_cell)

    df_in_out = df_in[["_SUBJECT_CLEAN", subj_col, visit_col, bm_col, pb_col, cd_col]].copy()
    df_in_out.rename(columns={
        subj_col:  "SUBJECT",
        visit_col: "VISIT",
        bm_col:    "BMBLASTS",
        pb_col:    "PBBLASTS",
        cd_col:    "CD123",
    }, inplace=True)
    df_in_out["SUBJECT"] = df_in_out["SUBJECT"].map(lambda x: None if pd.isna(x) else str(x).strip())

    out = {}
    for grp, subject_list in groups.items():
        mask = df_in_out["_SUBJECT_CLEAN"].isin(subject_list)
        part = df_in_out.loc[mask, ["SUBJECT", "VISIT", "BMBLASTS", "PBBLASTS", "CD123"]].copy()

        missing_subjects = [s for s in subject_list if s not in set(part["SUBJECT"].dropna())]
        if missing_subjects:
            filler = pd.DataFrame({
                "SUBJECT":  missing_subjects,
                "VISIT":    [pd.NA]*len(missing_subjects),
                "BMBLASTS": [pd.NA]*len(missing_subjects),
                "PBBLASTS": [pd.NA]*len(missing_subjects),
                "CD123":    [pd.NA]*len(missing_subjects),
            })
            part = pd.concat([part, filler], ignore_index=True)

        if "VISIT" in part.columns:
            part.sort_values(by=["SUBJECT", "VISIT"], inplace=True, kind="mergesort")
        else:
            part.sort_values(by=["SUBJECT"], inplace=True, kind="mergesort")

        out[grp] = part.reset_index(drop=True)
    return out

def compute_tp53_summary_exact(groups: dict, input_df: pd.DataFrame, subj_col_actual: str, tp53_col_actual: str) -> pd.DataFrame:
    """
    Build 'TP53 mutation' table with columns [Cohort, N patients, TP53 mutated, Percentage].
    - A subject counts as mutated iff ANY of their rows has TP53FINALRESULT == 'positive' (case/space-insensitive).
    - Subjects absent from input are treated as non-mutated (since no 'positive' evidence).
    """
    tmp = input_df[[subj_col_actual, tp53_col_actual]].copy()
    tmp["_SUBJECT_CLEAN"] = tmp[subj_col_actual].map(_clean_cell)

    def tp53_pos(v):
        if pd.isna(v):
            return False
        s = str(v).strip().lower()
        return s == "positive"

    tmp["__mut__"] = tmp[tp53_col_actual].map(tp53_pos)
    # subject-level mutation status (True if any row is positive)
    mut_map = tmp.groupby("_SUBJECT_CLEAN")["__mut__"].any().to_dict()

    rows = []
    for grp, subject_list in groups.items():
        n_pat = len(subject_list)
        mut_count = sum(1 for s in subject_list if mut_map.get(s, False))
        perc = round((mut_count / n_pat * 100.0), 2) if n_pat > 0 else 0.0
        rows.append({"Cohort": grp, "N patients": n_pat, "TP53 mutated": mut_count, "Percentage": perc})

    return pd.DataFrame(rows).sort_values(by="Cohort", kind="mergesort").reset_index(drop=True)

def write_excel_with_tp53(group_tables: dict, tp53_df: pd.DataFrame, out_path: str):
    used = set()
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for grp, df_pat in group_tables.items():
            sheet = sanitize_sheet_name(grp, used)
            df_pat.to_excel(writer, sheet_name=sheet, index=False)
        tp53_df.to_excel(writer, sheet_name="TP53 mutation", index=False)

# ---------- GUI ----------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        try:
            os.chdir(os.path.expanduser("~"))
        except Exception:
            pass

        self.title("Input/Dose → Excel by Treatment Group")
        self.geometry("800x360")
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
            messagebox.showerror("Missing input", "Please select a valid INPUT data file (with SUBJECT).")
            return
        if not dose_path or not os.path.isfile(dose_path):
            messagebox.showerror("Missing dose dict", "Please select a valid patient dose dictionary file.")
            return
        if not out_dir or not os.path.isdir(out_dir):
            messagebox.showerror("Missing output folder", "Please choose a valid output folder.")
            return
        if out_name == "":
            messagebox.showerror("Missing output name", "Please provide an output file name (e.g., dose_groups.xlsx).")
            return
        if not out_name.lower().endswith(".xlsx"):
            out_name += ".xlsx"
        out_path = os.path.join(out_dir, out_name)

        try:
            dose_df  = read_table(dose_path)
            input_df = read_table(in_path)

            # Dose dict columns (robust)
            dose_cols = match_columns(
                dose_df,
                {
                    "patientnumber": ["Patient Number"],
                    "treatmentgroup": ["Treatment Group"],
                },
            )
            col_patient = dose_cols["patientnumber"]
            col_group   = dose_cols["treatmentgroup"]

            # Groups → subjects (unique)
            groups = prepare_groups(dose_df, col_patient, col_group)

            # Input required columns (robust)
            input_cols = match_columns(
                input_df,
                {
                    "subject":  ["SUBJECT"],
                    "visit":    ["VISIT"],
                    "bmblasts": ["BMBLASTS", "BM BLASTS"],
                    "pbblasts": ["PBBLASTS", "PB BLASTS"],
                    "cd123":    ["CD123"],
                    "tp53":     ["TP53FINALRESULT"],  # exact field per your spec
                },
            )

            # Group sheets (SUBJECT, VISIT, BMBLASTS, PBBLASTS, CD123)
            group_tables = build_group_tables(input_df, input_cols, groups)
            if not group_tables:
                raise ValueError("No groups/patients found after cleaning. Check the dose dictionary content.")

            # TP53 mutation summary (positive vs negative)
            tp53_df = compute_tp53_summary_exact(groups, input_df, input_cols["subject"], input_cols["tp53"])

            # Write Excel
            write_excel_with_tp53(group_tables, tp53_df, out_path)

        except Exception as e:
            messagebox.showerror("Processing error", str(e))
            return

        messagebox.showinfo("Done", f"Excel written to:\n{friendly_path(out_path)}")

# ---------- entry ----------
if __name__ == "__main__":
    app = App()
    app.mainloop()
