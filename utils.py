import pandas as pd
import numpy as np
import io
import re


def col_letter_to_index(letter: str) -> int:
    letter = letter.strip().upper()
    result = 0
    for ch in letter:
        if not ch.isalpha():
            continue
        result = result * 26 + (ord(ch) - ord('A') + 1)
    return result - 1


def normalize_code(x):
    if pd.isna(x):
        return None
    if isinstance(x, (int, np.integer)):
        return str(int(x)).upper()
    if isinstance(x, (float, np.floating)):
        if np.isfinite(x) and float(x).is_integer():
            return str(int(x)).upper()
        return str(x).strip().upper()
    s = str(x).strip()
    if s.endswith('.0'):
        s = s[:-2]
    s = s.replace(' ', '')
    return s.upper()


def parse_price(v):
    if pd.isna(v):
        return np.nan
    if isinstance(v, (int, float, np.number)):
        return float(v)
    s = str(v).strip()
    s = s.replace("€", "").replace(" ", "")
    if "," in s and "." in s:
        s = s.replace(".", "")
        s = s.replace(",", ".")
    elif "," in s and "." not in s:
        s = s.replace(",", ".")
    s = re.sub(r"[^0-9\.\-]", "", s)
    try:
        return float(s)
    except:
        return np.nan


def load_price_list_by_letter(file_like, code_col_letter="C", price_col_letter="AK"):
    idx_code = col_letter_to_index(code_col_letter)
    idx_price = col_letter_to_index(price_col_letter)
    df = pd.read_excel(file_like, header=0)
    if df.shape[1] <= max(idx_code, idx_price):
        file_like.seek(0)
        df = pd.read_excel(file_like, header=None)
    code_series = df.iloc[:, idx_code]
    price_series = df.iloc[:, idx_price]
    out = pd.DataFrame({
        "code": code_series.apply(normalize_code),
        "list_price": price_series.apply(parse_price),
    })
    out = out.dropna(subset=["code"])
    out = out[~(out["code"].astype(str).str.strip() == "")]
    out = out.groupby("code", as_index=False).agg({"list_price": "first"})
    return out


def load_invoice_by_letter(file_like, code_col_letter="A", price_col_letter="D", source_name="fattura.xlsx"):
    idx_code = col_letter_to_index(code_col_letter)
    idx_price = col_letter_to_index(price_col_letter)
    df = pd.read_excel(file_like, header=0)
    if df.shape[1] <= max(idx_code, idx_price):
        file_like.seek(0)
        df = pd.read_excel(file_like, header=None)
    code_series = df.iloc[:, idx_code]
    price_series = df.iloc[:, idx_price]
    out = pd.DataFrame({
        "source_file": source_name,
        "code": code_series.apply(normalize_code),
        "invoice_price": price_series.apply(parse_price),
        "row_index": range(1, len(df) + 1),
    })
    out = out.dropna(subset=["code"])
    out = out[~(out["code"].astype(str).str.strip() == "")]
    return out


def match_invoices_to_pricelist(inv_df, price_df, tol_abs=0.01, tol_pct=0.0):
    merged = inv_df.merge(price_df, on="code", how="left")
    def status_row(r):
        lp = r.get("list_price", np.nan)
        ip = r.get("invoice_price", np.nan)
        if pd.isna(lp):
            return "CODICE_NON_TROVATO"
        if pd.isna(ip):
            return "PREZZO_NON_VALIDO"
        abs_diff = abs(ip - lp)
        pct_diff = abs_diff / lp if lp not in (0, np.nan) else np.nan
        ok_abs = abs_diff <= tol_abs + 1e-12
        ok_pct = True if (tol_pct is None or tol_pct <= 0) else (pct_diff <= tol_pct / 100.0)
        if ok_abs and ok_pct:
            return "OK"
        else:
            return "PREZZO_DIVERSO"
    merged["status"] = merged.apply(status_row, axis=1)
    merged["delta_abs"] = merged["invoice_price"] - merged["list_price"]
    merged["delta_pct"] = merged["delta_abs"] / merged["list_price"]
    merged["price_match"] = merged["status"].eq("OK")
    cols = ["source_file","row_index","code","invoice_price","list_price","delta_abs","delta_pct","status","price_match"]
    merged = merged[cols]
    status_order = {"PREZZO_DIVERSO": 0, "CODICE_NON_TROVATO": 1, "PREZZO_NON_VALIDO": 2, "OK": 3}
    merged["__order"] = merged["status"].map(status_order).fillna(9)
    merged = merged.sort_values(["__order","source_file","code","row_index"]).drop(columns="__order")
    return merged


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Report")
        wb = writer.book
        ws = writer.sheets["Report"]
        cols = list(df.columns)
        if "delta_pct" in cols:
            col_idx = cols.index("delta_pct")
            ws.set_column(col_idx, col_idx, 12, wb.add_format({"num_format": "0.00%"}))
        for name in ["invoice_price","list_price","delta_abs"]:
            if name in cols:
                i = cols.index(name)
                ws.set_column(i, i, 14, wb.add_format({"num_format": "\u20ac #,##0.00"}))
        ws.set_zoom(110)
    return bio.getvalue()
