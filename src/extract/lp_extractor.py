"""
Loading Plan Extractor (06-Loading Plan)
Extracts required loading dates from three sources: KS, IDN Export, IDN Domestic.
"""
import pandas as pd
import re
from datetime import timedelta


def extract_lp_ks(file_path: str, data_date: str) -> pd.DataFrame:
    """
    Extract KS Loading Plan with invoice number splitting.

    Args:
        file_path: Path to LoadingPlan Excel file
        data_date: Data extraction date string (e.g. "2026-05-08")

    Returns:
        DataFrame with columns: so, loading_date, load_mt, source
    """
    df = pd.read_excel(file_path, sheet_name="Loading Plan")
    df.columns = df.columns.str.strip()

    # Identify columns by position (A=0, D=3, E=4, F=5, H=7)
    cols = df.columns.tolist()
    inv_col = cols[0]       # A: Invoice No.
    gp20_col = cols[3]      # D: 20GP
    gp40_col = cols[4]      # E: 40GP
    hq40_col = cols[5]      # F: 40HQ
    loading_col = cols[7]   # H: Loading Date

    # Time filter: >= data_date - 3 days
    cutoff = pd.to_datetime(data_date) - timedelta(days=3)
    df["_loading_date"] = pd.to_datetime(df[loading_col], errors="coerce")
    df = df[df["_loading_date"] >= cutoff].copy()

    # Calculate tonnage: 20GP×15 + 40GP×24.5 + 40HQ×24.5
    gp20 = pd.to_numeric(df[gp20_col], errors="coerce").fillna(0)
    gp40 = pd.to_numeric(df[gp40_col], errors="coerce").fillna(0)
    hq40 = pd.to_numeric(df[hq40_col], errors="coerce").fillna(0)
    df["_total_mt"] = gp20 * 15 + gp40 * 24.5 + hq40 * 24.5

    # Split invoice numbers and expand
    records = []
    for _, row in df.iterrows():
        inv_raw = str(row[inv_col]).strip()
        loading_date = row["_loading_date"]
        total_mt = row["_total_mt"]

        if pd.isna(loading_date) or total_mt <= 0:
            continue

        so_list = _split_ks_invoice(inv_raw)
        if not so_list:
            continue

        mt_per_so = total_mt / len(so_list)
        for so in so_list:
            records.append({
                "so": so,
                "loading_date": loading_date,
                "load_mt": mt_per_so,
                "source": "KS_LP"
            })

    result = pd.DataFrame(records)
    if result.empty:
        return pd.DataFrame(columns=["so", "loading_date", "load_mt", "source"])

    # Aggregate by SO (take earliest loading date if multiple entries)
    agg = result.groupby("so").agg(
        loading_date=("loading_date", "min"),
        load_mt=("load_mt", "sum"),
        source=("source", "first")
    ).reset_index()

    return agg


def extract_lp_idn_export(file_path: str, data_date: str) -> pd.DataFrame:
    """
    Extract IDN Export Loading Plan from ORDER OUTSTANDING sheet.

    Args:
        file_path: Path to Schedule Planning Dispatch Excel file
        data_date: Data extraction date string

    Returns:
        DataFrame with columns: so, loading_date, load_mt, source, cluster
    """
    df = pd.read_excel(file_path, sheet_name="ORDER OUTSTANDING ", header=1)
    df.columns = df.columns.str.strip()

    # Key columns by position: C=SC No.(idx2), E=Region(idx4), J=Cont Qty(idx9), K=Cont Size(idx10), S=ELD(idx18)
    cols = df.columns.tolist()
    so_col = cols[2]        # C: SC No.
    region_col = cols[4]    # E: Region
    qty_col = cols[9]       # J: Cont Qty
    size_col = cols[10]     # K: Cont Size
    eld_col = cols[18]      # S: ELD

    result = pd.DataFrame()
    result["so"] = df[so_col].apply(_to_so_str)
    result["cluster"] = df[region_col].astype(str).str.strip()
    result["loading_date"] = pd.to_datetime(df[eld_col], errors="coerce")

    # Tonnage: Cont Qty × MT per size
    cont_qty = pd.to_numeric(df[qty_col], errors="coerce").fillna(0)
    cont_size = df[size_col].astype(str).str.strip()
    mt_per_cont = cont_size.apply(_map_container_mt)
    result["load_mt"] = cont_qty * mt_per_cont
    result["source"] = "IDN_Export"

    # Time filter
    cutoff = pd.to_datetime(data_date) - timedelta(days=3)
    result = result[result["loading_date"] >= cutoff].copy()

    # Remove invalid
    result = result[result["so"].str.match(r"^\d+$", na=False)].copy()
    result = result[result["load_mt"] > 0].copy()

    # Aggregate by SO
    agg = result.groupby("so").agg(
        loading_date=("loading_date", "min"),
        load_mt=("load_mt", "sum"),
        source=("source", "first"),
        cluster=("cluster", "first")
    ).reset_index()

    return agg


def extract_lp_idn_domestic(file_path: str, data_date: str) -> pd.DataFrame:
    """
    Extract IDN Domestic Loading Plan from Order List sheet.

    Args:
        file_path: Path to NEW DOMESTIC TRACKING Excel file
        data_date: Data extraction date string

    Returns:
        DataFrame with columns: so, loading_date, load_mt, source
    """
    df = pd.read_excel(file_path, sheet_name="Order List", header=1)
    df.columns = df.columns.str.strip()

    # Key columns: D=INV NO.(idx3), G=ELD(idx6), P=Weight(idx15)
    cols = df.columns.tolist()
    inv_col = cols[3]       # D: INV NO.
    eld_col = cols[6]       # G: ELD
    weight_col = cols[15]   # P: Weight (already MT)

    # Time filter
    cutoff = pd.to_datetime(data_date) - timedelta(days=3)
    df["_eld"] = pd.to_datetime(df[eld_col], errors="coerce")
    df = df[df["_eld"] >= cutoff].copy()

    # Split INV NO. and expand
    records = []
    for _, row in df.iterrows():
        inv_raw = str(row[inv_col]).strip()
        eld = row["_eld"]
        weight = pd.to_numeric(row[weight_col], errors="coerce")

        if pd.isna(eld) or pd.isna(weight) or weight <= 0:
            continue

        so_list = _split_idn_domestic_inv(inv_raw)
        if not so_list:
            continue

        mt_per_so = weight / len(so_list)
        for so in so_list:
            records.append({
                "so": so,
                "loading_date": eld,
                "load_mt": mt_per_so,
                "source": "IDN_Domestic"
            })

    result = pd.DataFrame(records)
    if result.empty:
        return pd.DataFrame(columns=["so", "loading_date", "load_mt", "source"])

    # Aggregate by SO
    agg = result.groupby("so").agg(
        loading_date=("loading_date", "min"),
        load_mt=("load_mt", "sum"),
        source=("source", "first")
    ).reset_index()

    return agg


# --- Splitting helpers ---

def _split_ks_invoice(inv: str) -> list:
    """
    Split KS invoice number into SO numbers.
    - Underscore '_' = multi-SO separator
    - Dash '-' = container sequence (strip)
    """
    if not inv or inv.lower() == "nan":
        return []

    # Strip dash suffix first (container sequence like -1, -2)
    inv_clean = re.sub(r"-\d+$", "", inv.strip())

    # If no underscore, it's a single SO
    if "_" not in inv_clean:
        if re.match(r"^10\d{8}$", inv_clean):
            return [inv_clean]
        # Try parsing as pure number
        try:
            val = str(int(float(inv_clean)))
            if re.match(r"^10\d{8}$", val):
                return [val]
        except (ValueError, TypeError):
            pass
        return []

    # Split by underscore
    parts = inv_clean.split("_")

    # Find the longest part as prefix source
    longest = max(parts, key=len)

    so_list = []
    for part in parts:
        clean = part.strip()
        if not clean:
            continue
        # Strip any remaining dash suffix within split parts
        clean = re.sub(r"-\d+$", "", clean)
        if not clean:
            continue
        # Prefix completion for short numbers
        if len(clean) < 10 and len(longest) == 10:
            needed = 10 - len(clean)
            clean = longest[:needed] + clean
        so_list.append(clean)

    # Validate: all should be 10-digit numbers starting with 10
    valid = [s for s in so_list if re.match(r"^10\d{8}$", s)]
    return valid if valid else []


def _split_idn_domestic_inv(inv: str) -> list:
    """
    Split IDN Domestic INV NO. into SO numbers.
    - Non-10 prefix (e.g. LMIDSAM*) → remove
    - Comma ',' or '&' = multi-SO separator
    - Dash '-' = batch/trip number (strip suffix)
    """
    if not inv or inv.lower() == "nan":
        return []

    # Filter: only process 10-prefixed orders
    if not re.search(r"10\d{7}", inv):
        return []

    # Split by comma or &
    parts = re.split(r"[,&]", inv)

    so_list = []
    for part in parts:
        clean = part.strip()
        # Remove non-breaking spaces
        clean = clean.replace("\xa0", "").strip()
        # Strip dash suffix
        clean = re.sub(r"-\d+$", "", clean)
        if not clean:
            continue
        # Must be a valid SO number
        if re.match(r"^10\d{8}$", clean):
            so_list.append(clean)

    return so_list


def _to_so_str(value) -> str:
    """Convert SO value (often float) to clean string."""
    if pd.isna(value):
        return ""
    try:
        return str(int(float(value)))
    except (ValueError, TypeError):
        return str(value).strip()


def _map_container_mt(size_str: str) -> float:
    """Map container size to MT."""
    s = size_str.lower().replace("'", "").replace("'", "")
    if "40" in s or "hc" in s or "hq" in s:
        return 24.5
    elif "20" in s:
        return 15.0
    return 15.0  # default


def combine_loading_plans(lp_ks: pd.DataFrame, lp_idn_export: pd.DataFrame, lp_idn_domestic: pd.DataFrame) -> pd.DataFrame:
    """
    Combine all three loading plan sources into one unified DataFrame.

    Returns:
        DataFrame with columns: so, loading_date, load_mt, source
    """
    # Standardize columns
    for df in [lp_ks, lp_idn_export, lp_idn_domestic]:
        if "cluster" in df.columns:
            df.drop(columns=["cluster"], inplace=True, errors="ignore")

    combined = pd.concat([lp_ks, lp_idn_export, lp_idn_domestic], ignore_index=True)

    if combined.empty:
        return pd.DataFrame(columns=["so", "loading_date", "load_mt", "source"])

    # If same SO appears in multiple sources, take earliest loading_date
    agg = combined.groupby("so").agg(
        loading_date=("loading_date", "min"),
        load_mt=("load_mt", "sum"),
        source=("source", "first")
    ).reset_index()

    return agg
