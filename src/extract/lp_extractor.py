"""
Loading Plan Extractor (06-Loading Plan)

Builds clean demand-line records from KS and IDN loading plan source workbooks.
The extractor intentionally preserves unconfirmed dates and parse exceptions so
they can be reviewed as business risks instead of being silently filtered out.
"""
import re
from pathlib import Path

import pandas as pd


MT_20GP = 14.5
MT_40GP = 24.5
STANDARD_COLUMNS = [
    "plant",
    "lp_source",
    "source_file",
    "source_sheet",
    "source_row",
    "invoice_no_raw",
    "so",
    "so_parse_status",
    "so_parse_note",
    "loading_date_raw",
    "loading_date",
    "loading_date_status",
    "load_mt",
    "source_mt",
    "exclude_from_current_invoice",
    "exclude_reason",
    "split_count",
]


def extract_lp_ks(file_path: str, data_date: str = "") -> pd.DataFrame:
    """Extract KS loading plan demand lines from the raw Loading plan sheet."""
    source_sheet = "Loading plan"
    df = pd.read_excel(file_path, sheet_name=source_sheet)
    df.columns = _dedupe_columns([str(c).strip() for c in df.columns])

    inv_col = _find_column(df, "Invoice No")
    gp20_col = _find_column(df, "20GP")
    gp40_col = _find_column(df, "40GP")
    hq40_col = _find_column(df, "40HQ")
    loading_col = _find_column(df, "Loading")
    source_mt_col = _find_column(df, "MT", required=False)
    exclude_col = df.columns[20] if len(df.columns) >= 21 else None

    records = []
    for idx, row in df.iterrows():
        invoice_raw = _clean_text(row.get(inv_col))
        loading_raw = row.get(loading_col)
        loading_date, loading_status = _parse_loading_date(loading_raw)

        gp20 = _to_number(row.get(gp20_col))
        gp40 = _to_number(row.get(gp40_col))
        hq40 = _to_number(row.get(hq40_col))
        load_mt = gp20 * MT_20GP + gp40 * MT_40GP + hq40 * MT_40GP
        source_mt = _to_number(row.get(source_mt_col)) if source_mt_col else 0.0

        exclude_reason = _clean_text(row.get(exclude_col)) if exclude_col else ""
        excluded = bool(exclude_reason)

        if not invoice_raw and pd.isna(loading_raw) and load_mt == 0 and source_mt == 0:
            continue

        so_list, parse_status, parse_note = _split_ks_invoice(invoice_raw)
        records.extend(
            _expand_record(
                row=row,
                base={
                    "plant": "KS",
                    "lp_source": "KS_LP",
                    "source_file": Path(file_path).name,
                    "source_sheet": source_sheet,
                    "source_row": int(idx) + 2,
                    "invoice_no_raw": invoice_raw,
                    "loading_date_raw": _clean_text(loading_raw),
                    "loading_date": loading_date,
                    "loading_date_status": loading_status,
                    "load_mt": load_mt,
                    "source_mt": source_mt,
                    "exclude_from_current_invoice": excluded,
                    "exclude_reason": exclude_reason,
                },
                so_list=so_list,
                parse_status=parse_status,
                parse_note=parse_note,
            )
        )

    return _standardize_output(records)


def extract_lp_idn_export(file_path: str, data_date: str = "") -> pd.DataFrame:
    """Extract IDN export loading plan demand lines from ORDER OUTSTANDING."""
    source_sheet = "ORDER OUTSTANDING "
    df = pd.read_excel(file_path, sheet_name=source_sheet, header=1)
    df.columns = _dedupe_columns([str(c).strip() for c in df.columns])

    so_col = _find_column(df, "SC No.")
    qty_col = _find_column(df, "Cont Qty")
    size_col = _find_column(df, "Cont Size")
    loading_col = _find_column(df, "ELD")
    source_mt_col = _find_column(df, "Rough Ton", required=False)

    records = []
    for idx, row in df.iterrows():
        customer = _clean_text(row.get(_find_column(df, "Customers name", required=False)))
        so_raw = _clean_text(row.get(so_col))
        if not so_raw and not customer:
            continue

        loading_raw = row.get(loading_col)
        loading_date, loading_status = _parse_loading_date(loading_raw)
        qty = _to_number(row.get(qty_col))
        size = _clean_text(row.get(size_col))
        load_mt = qty * _container_size_mt(size)
        source_mt = _to_number(row.get(source_mt_col)) if source_mt_col else 0.0

        so = _to_code_str(so_raw)
        if re.fullmatch(r"10\d{8}", so):
            so_list, parse_status, parse_note = [so], "Parsed", "SC No. direct"
        else:
            so_list, parse_status, parse_note = [], "Parse Failed", "SC No. is not a standard SO"

        records.extend(
            _expand_record(
                row=row,
                base={
                    "plant": "IDN",
                    "lp_source": "IDN_Export",
                    "source_file": Path(file_path).name,
                    "source_sheet": source_sheet,
                    "source_row": int(idx) + 3,
                    "invoice_no_raw": so_raw,
                    "loading_date_raw": _clean_text(loading_raw),
                    "loading_date": loading_date,
                    "loading_date_status": loading_status,
                    "load_mt": load_mt,
                    "source_mt": source_mt,
                    "exclude_from_current_invoice": False,
                    "exclude_reason": "",
                },
                so_list=so_list,
                parse_status=parse_status,
                parse_note=parse_note,
            )
        )

    return _standardize_output(records)


def extract_lp_idn_domestic(file_path: str, data_date: str = "") -> pd.DataFrame:
    """Extract IDN domestic loading plan lines, preferring SC NO. over INV NO."""
    source_sheet = "Order List"
    df = pd.read_excel(file_path, sheet_name=source_sheet, header=1)
    df.columns = _dedupe_columns([str(c).strip() for c in df.columns])

    inv_col = _find_column(df, "INV NO.")
    sc_col = _find_column(df, "SC NO.", required=False)
    loading_col = _find_column(df, "ELD")
    weight_col = _find_column(df, "Weight")

    records = []
    for idx, row in df.iterrows():
        invoice_raw = _clean_text(row.get(inv_col))
        sc_raw = _clean_text(row.get(sc_col)) if sc_col else ""
        customer_col = _find_column(df, "CUSTOMER NAME", required=False)
        customer = _clean_text(row.get(customer_col)) if customer_col else ""
        if not invoice_raw and not sc_raw and not customer:
            continue

        loading_raw = row.get(loading_col)
        loading_date, loading_status = _parse_loading_date(loading_raw)
        load_mt = _to_number(row.get(weight_col))
        source_mt = load_mt

        sc_so = _to_code_str(sc_raw)
        if re.fullmatch(r"10\d{8}", sc_so):
            so_list, parse_status, parse_note = [sc_so], "Parsed", "SC NO. direct"
        else:
            so_list, parse_status, parse_note = _split_idn_domestic_inv(invoice_raw)

        records.extend(
            _expand_record(
                row=row,
                base={
                    "plant": "IDN",
                    "lp_source": "IDN_Domestic",
                    "source_file": Path(file_path).name,
                    "source_sheet": source_sheet,
                    "source_row": int(idx) + 3,
                    "invoice_no_raw": invoice_raw,
                    "loading_date_raw": _clean_text(loading_raw),
                    "loading_date": loading_date,
                    "loading_date_status": loading_status,
                    "load_mt": load_mt,
                    "source_mt": source_mt,
                    "exclude_from_current_invoice": False,
                    "exclude_reason": "",
                },
                so_list=so_list,
                parse_status=parse_status,
                parse_note=parse_note,
            )
        )

    return _standardize_output(records)


def combine_loading_plan_lines(*frames: pd.DataFrame) -> pd.DataFrame:
    """Combine clean demand-line outputs without aggregating away dates."""
    non_empty = [f for f in frames if f is not None and not f.empty]
    if not non_empty:
        return _standardize_output([])
    combined = pd.concat(non_empty, ignore_index=True, sort=False)
    for col in STANDARD_COLUMNS:
        if col not in combined.columns:
            combined[col] = pd.NA
    combined["exclude_from_current_invoice"] = combined["exclude_from_current_invoice"].fillna(False).astype(bool)
    combined["load_mt"] = pd.to_numeric(combined["load_mt"], errors="coerce").fillna(0)
    combined["source_mt"] = pd.to_numeric(combined["source_mt"], errors="coerce").fillna(0)
    return combined


def aggregate_loading_plan_for_master(clean_lines: pd.DataFrame) -> pd.DataFrame:
    """
    Build a SO-level view for the existing SO Master join.

    This intentionally keeps unconfirmed loading demand as LP records while only
    using valid dates for the legacy gap date.
    """
    if clean_lines is None or clean_lines.empty:
        return pd.DataFrame(columns=["so", "loading_date", "load_mt", "source"])

    main = clean_lines[
        (~clean_lines["exclude_from_current_invoice"])
        & (clean_lines["so_parse_status"] == "Parsed")
        & clean_lines["so"].notna()
        & (clean_lines["so"].astype(str).str.strip() != "")
    ].copy()
    if main.empty:
        return pd.DataFrame(columns=["so", "loading_date", "load_mt", "source"])

    main["valid_load_mt"] = main["load_mt"].where(main["loading_date_status"] == "Valid Date", 0)
    main["unconfirmed_load_mt"] = main["load_mt"].where(main["loading_date_status"] != "Valid Date", 0)

    agg = main.groupby("so").agg(
        loading_date=("loading_date", "min"),
        load_mt=("load_mt", "sum"),
        lp_valid_mt=("valid_load_mt", "sum"),
        lp_unconfirmed_mt=("unconfirmed_load_mt", "sum"),
        lp_line_count=("so", "size"),
        source=("lp_source", lambda x: ", ".join(sorted(set(map(str, x))))),
        lp_date_status=("loading_date_status", _status_label),
    ).reset_index()
    agg["lp_has_record"] = True
    return agg


def combine_loading_plans(lp_ks: pd.DataFrame, lp_idn_export: pd.DataFrame, lp_idn_domestic: pd.DataFrame) -> pd.DataFrame:
    """Backward-compatible wrapper returning a SO-level loading plan view."""
    lines = combine_loading_plan_lines(lp_ks, lp_idn_export, lp_idn_domestic)
    agg = aggregate_loading_plan_for_master(lines)
    agg.attrs["clean_loading_plan"] = lines
    return agg


def _expand_record(row, base: dict, so_list: list[str], parse_status: str, parse_note: str) -> list[dict]:
    split_count = len(so_list)
    load_per_so = base["load_mt"] / split_count if split_count else base["load_mt"]
    source_per_so = base["source_mt"] / split_count if split_count else base["source_mt"]
    targets = so_list or [""]
    records = []
    for so in targets:
        record = dict(base)
        record["so"] = so
        record["so_parse_status"] = parse_status
        record["so_parse_note"] = parse_note
        record["split_count"] = split_count
        record["load_mt"] = load_per_so
        record["source_mt"] = source_per_so
        for col, val in row.items():
            record[f"orig_{col}"] = val
        records.append(record)
    return records


def _split_ks_invoice(inv: str) -> tuple[list[str], str, str]:
    """Split KS Invoice No. into SOs while preserving parse failures."""
    inv = _clean_text(inv)
    if not inv:
        return [], "Parse Failed", "Blank Invoice No"
    if inv.upper().startswith("LM"):
        return [], "Non-SC", "Non-standard LM invoice"

    parts = re.split(r"_", inv)
    longest = max(parts, key=len) if parts else inv
    longest_digits = _first_10_digit(longest)
    so_list = []
    for part in parts:
        clean = _clean_text(part)
        clean = re.sub(r"-\s*\d+\s*~\s*\d+\s*$", "", clean)
        clean = re.sub(r"-\s*\d+\s*$", "", clean)
        if re.fullmatch(r"10\d{8}", clean):
            so_list.append(clean)
            continue
        if clean.isdigit() and len(clean) < 10 and longest_digits:
            needed = 10 - len(clean)
            candidate = longest_digits[:needed] + clean
            if re.fullmatch(r"10\d{8}", candidate):
                so_list.append(candidate)
                continue
        candidate = _first_10_digit(clean)
        if candidate:
            so_list.append(candidate)

    so_list = list(dict.fromkeys(so_list))
    if so_list:
        return so_list, "Parsed", "Invoice No parsed"
    return [], "Parse Failed", "No standard SO found in Invoice No"


def _split_idn_domestic_inv(inv: str) -> tuple[list[str], str, str]:
    """Split IDN Domestic INV NO. into SOs."""
    inv = _clean_text(inv)
    if not inv:
        return [], "Parse Failed", "Blank INV NO."
    if inv.upper().startswith("LM"):
        return [], "Non-SC", "Non-standard LM invoice"

    parts = re.split(r"[,&]", inv)
    so_list = []
    for part in parts:
        clean = _clean_text(part)
        clean = clean.replace("\u00a0", "").replace("\u2002", "").strip()
        clean = re.sub(r"-\s*\d+\s*~\s*\d+\s*$", "", clean)
        clean = re.sub(r"-\s*\d+\s*$", "", clean)
        candidate = _first_10_digit(clean)
        if candidate:
            so_list.append(candidate)

    so_list = list(dict.fromkeys(so_list))
    if so_list:
        return so_list, "Parsed", "INV NO. parsed"
    return [], "Parse Failed", "No standard SO found in INV NO."


def _parse_loading_date(value) -> tuple[pd.Timestamp, str]:
    if pd.isna(value):
        return pd.NaT, "Blank"
    text = _clean_text(value)
    if not text:
        return pd.NaT, "Blank"
    parsed = pd.to_datetime(value, errors="coerce")
    if pd.notna(parsed):
        return parsed, "Valid Date"
    upper = text.upper()
    if "TBA" in upper or "PENDING" in upper:
        return pd.NaT, "TBA"
    if re.search(r"(月|JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)", upper):
        return pd.NaT, "Text Month"
    return pd.NaT, "Invalid Text"


def _container_size_mt(size_str: str) -> float:
    size = _clean_text(size_str).lower().replace("'", "")
    if "20" in size:
        return MT_20GP
    if "40" in size or "hc" in size or "hq" in size:
        return MT_40GP
    return 0.0


def _status_label(values: pd.Series) -> str:
    statuses = [str(v) for v in values.dropna().unique().tolist() if str(v)]
    if not statuses:
        return ""
    if len(statuses) == 1:
        return statuses[0]
    if "Valid Date" in statuses:
        return "Mixed"
    return "Unconfirmed"


def _standardize_output(records: list[dict]) -> pd.DataFrame:
    df = pd.DataFrame(records)
    if df.empty:
        return pd.DataFrame(columns=STANDARD_COLUMNS)
    for col in STANDARD_COLUMNS:
        if col not in df.columns:
            df[col] = pd.NA
    ordered = STANDARD_COLUMNS + [c for c in df.columns if c not in STANDARD_COLUMNS]
    return df[ordered]


def _find_column(df: pd.DataFrame, name: str, required: bool = True):
    target = _normalize(name)
    for col in df.columns:
        if _normalize(col) == target:
            return col
    for col in df.columns:
        if target in _normalize(col):
            return col
    if required:
        raise ValueError(f"Column '{name}' not found. Available: {list(df.columns)}")
    return None


def _normalize(value) -> str:
    return " ".join(str(value).strip().lower().split())


def _dedupe_columns(columns: list[str]) -> list[str]:
    counts = {}
    result = []
    for col in columns:
        if col not in counts:
            counts[col] = 0
            result.append(col)
        else:
            counts[col] += 1
            result.append(f"{col}.{counts[col]}")
    return result


def _to_number(value) -> float:
    number = pd.to_numeric(value, errors="coerce")
    if pd.isna(number):
        return 0.0
    return float(number)


def _clean_text(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).replace("\xa0", " ").replace("\u2002", " ").strip()


def _to_code_str(value) -> str:
    text = _clean_text(value)
    if not text:
        return ""
    try:
        number = float(text)
        if number.is_integer():
            return str(int(number))
    except (ValueError, TypeError):
        pass
    if text.endswith(".0"):
        return text[:-2]
    return text


def _first_10_digit(value: str) -> str:
    match = re.search(r"10\d{8}", _clean_text(value))
    return match.group(0) if match else ""
