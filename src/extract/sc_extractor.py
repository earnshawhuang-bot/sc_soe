"""
SC Extractor - Sales Order Baseline (01-SC)
Extracts and classifies orders from the SC master table.
"""
import pandas as pd


ADJUSTMENT_FACTOR = 0.975
BASELINE_TYPES = [
    "Carry Over Stock",
    "Carry Over Unproduced",
    "Fresh Order This Month",
]
ADJUSTED_TYPES = [
    "Carry Over Unproduced",
    "Fresh Order This Month",
    "Fresh Order Next Month",
]


def extract_sc(file_path: str, month: str, customer_mapping_path: str = None, data_date: str = None) -> pd.DataFrame:
    """
    Extract SC baseline orders.

    Args:
        file_path: Path to Order tracking Excel file
        month: Current month string e.g. "2026-05"
        customer_mapping_path: Optional path to dim_cc_region.xlsx
        data_date: Run data date, used for prior-month SC delivery pre-allocation

    Returns:
        SO-level baseline DataFrame. sc_vol_mt is adjusted baseline volume.
    """
    # Read the specific sheet
    df = pd.read_excel(file_path, sheet_name="Order Status - Main")

    # Standardize column names (strip whitespace)
    df.columns = df.columns.str.strip()

    # Filter: Item = "main"
    item_col = _find_column(df, "Item")
    df = df[df[item_col].astype(str).str.strip().str.lower() == "main"].copy()

    # Locate SC columns by stable headers. The file owner may insert columns
    # month to month, so fixed Excel positions are not reliable.
    so_col = _find_column(df, "SC NO")
    plant_col = _find_column(df, "Supply from")
    cluster_col = _find_column(df, "Cluster")
    vol_col = _find_column(df, "SC Vol.-MT")
    end_customer_col = _find_column(df, "End User Cust ID")
    status_col = _find_column(df, "Order status")
    q_col = _find_column(df, "Carryover")  # null = this month's billing
    ao_col = _find_column(df, "carryover/fresh production")
    release_col = _find_column(df, "RELEASE DATE")
    loading_col = _find_column(df, "Loading date")
    delivery_col = _find_column(df, "Delivery  PCS", fallback="Delivery PCS")

    # Build result DataFrame
    result = pd.DataFrame()
    result["so"] = df[so_col].apply(_to_so_str)
    result["plant"] = df[plant_col].apply(_map_plant)
    result["cluster"] = df[cluster_col].astype(str).str.strip()
    result["end_customer_id"] = df[end_customer_col].apply(_to_code_str)
    result["raw_sc_vol_mt"] = pd.to_numeric(df[vol_col], errors="coerce").fillna(0)
    result["loading_date_sc"] = pd.to_datetime(df[loading_col], errors="coerce")
    result["delivery_pcs"] = pd.to_numeric(df[delivery_col], errors="coerce")

    # Classification inputs
    order_status  = df[status_col].astype(str).str.strip()
    q_col_values  = df[q_col]
    ao_col_values = df[ao_col].astype(str).str.strip().str.lower()
    release_dates = pd.to_datetime(df[release_col], errors="coerce")

    # Q null = this month's billing order
    q_is_null = q_col_values.isna()

    # Current month check
    target_year  = int(month.split("-")[0])
    target_month = int(month.split("-")[1])
    is_current_month = (release_dates.dt.year == target_year) & (release_dates.dt.month == target_month)

    # Four order types (independently verified, MECE)
    result["order_type"] = "Unknown"

    # Type 1: Carry Over Stock  (K='Carryover' AND Q null)
    mask_cos = (order_status == "Carryover") & q_is_null
    result.loc[mask_cos, "order_type"] = "Carry Over Stock"

    # Type 2: Carry Over Unproduced  (AO='last month' AND Q null)
    mask_coup = (ao_col_values == "last month") & q_is_null
    result.loc[mask_coup, "order_type"] = "Carry Over Unproduced"

    # Type 3: Fresh Order This Month  (Release Date = current month AND Q null)
    mask_fresh = is_current_month & q_is_null
    result.loc[mask_fresh, "order_type"] = "Fresh Order This Month"

    # Type 4: Fresh Order Next Month  (Release Date = current month AND Q not null → excluded)
    mask_next = is_current_month & ~q_is_null
    result.loc[mask_next, "order_type"] = "Fresh Order Next Month"

    result["adjustment_factor"] = 1.0
    result.loc[result["order_type"].isin(ADJUSTED_TYPES), "adjustment_factor"] = ADJUSTMENT_FACTOR
    result["adjusted_sc_vol_mt"] = result["raw_sc_vol_mt"] * result["adjustment_factor"]

    # Baseline = Type 1 + 2 + 3 (exclude Type 4 and Unknown)
    result["in_baseline"] = result["order_type"].isin([
        "Carry Over Stock", "Carry Over Unproduced", "Fresh Order This Month"
    ])

    result["customer_mapping_matched"] = True
    if customer_mapping_path:
        customer_keys = _load_customer_keys(customer_mapping_path)
        result["customer_mapping_matched"] = result["end_customer_id"].isin(customer_keys)

    # Remove rows with invalid SO
    result = result[result["so"].str.match(r"^\d+$", na=False)].copy()

    unmatched_customer = result[~result["customer_mapping_matched"]].copy()
    valid_mapped = result[result["customer_mapping_matched"]].copy()
    valid_mapped["sc_prior_delivery_eligible"] = _prior_delivery_mask(valid_mapped, data_date)
    prior_delivery = _build_prior_delivery(valid_mapped, data_date)
    fresh_next = valid_mapped[valid_mapped["order_type"] == "Fresh Order Next Month"].copy()
    unknown = valid_mapped[valid_mapped["order_type"] == "Unknown"].copy()
    baseline_rows = valid_mapped[valid_mapped["in_baseline"]].copy()

    type_pivot = baseline_rows.pivot_table(
        index="so",
        columns="order_type",
        values="adjusted_sc_vol_mt",
        aggfunc="sum",
        fill_value=0,
    )
    type_pivot = type_pivot.reindex(columns=BASELINE_TYPES, fill_value=0).rename(
        columns={
            "Carry Over Stock": "carry_over_stock_mt",
            "Carry Over Unproduced": "carry_over_unproduced_mt",
            "Fresh Order This Month": "fresh_this_month_mt",
        }
    )

    raw_type_pivot = baseline_rows.pivot_table(
        index="so",
        columns="order_type",
        values="raw_sc_vol_mt",
        aggfunc="sum",
        fill_value=0,
    )
    raw_type_pivot = raw_type_pivot.reindex(columns=BASELINE_TYPES, fill_value=0).rename(
        columns={
            "Carry Over Stock": "raw_carry_over_stock_mt",
            "Carry Over Unproduced": "raw_carry_over_unproduced_mt",
            "Fresh Order This Month": "raw_fresh_this_month_mt",
        }
    )

    # Aggregate by SO after row-level classification and baseline filtering.
    agg = baseline_rows.groupby("so").agg(
        plant=("plant", "first"),
        cluster=("cluster", "first"),
        sc_vol_mt=("adjusted_sc_vol_mt", "sum"),
        raw_sc_vol_mt=("raw_sc_vol_mt", "sum"),
        loading_date_sc=("loading_date_sc", "first"),
        order_type=("order_type", _order_type_label),
        in_baseline=("in_baseline", "max"),
        sc_row_count=("so", "size"),
    ).reset_index()
    agg = agg.merge(type_pivot.reset_index(), on="so", how="left")
    agg = agg.merge(raw_type_pivot.reset_index(), on="so", how="left")
    qty_cols = [
        "carry_over_stock_mt", "carry_over_unproduced_mt", "fresh_this_month_mt",
        "raw_carry_over_stock_mt", "raw_carry_over_unproduced_mt", "raw_fresh_this_month_mt",
    ]
    for col in qty_cols:
        agg[col] = agg[col].fillna(0)
    agg = agg.merge(prior_delivery, on="so", how="left")
    agg["raw_sc_prior_delivery_mt"] = agg["raw_sc_prior_delivery_mt"].fillna(0)

    agg.attrs["sc_row_detail"] = _audit_frame(valid_mapped)
    agg.attrs["unmatched_customer"] = _audit_frame(unmatched_customer)
    agg.attrs["fresh_next_month"] = _audit_frame(fresh_next)
    agg.attrs["unknown_type"] = _audit_frame(unknown)
    agg.attrs["sc_prior_delivery"] = _audit_frame(valid_mapped[valid_mapped["sc_prior_delivery_eligible"]])

    return agg


def _build_prior_delivery(df: pd.DataFrame, data_date: str) -> pd.DataFrame:
    """Build SO-level SC prior-month delivery pre-allocation."""
    prior = df[_prior_delivery_mask(df, data_date)].copy()
    if prior.empty:
        return pd.DataFrame(columns=["so", "raw_sc_prior_delivery_mt"])

    return prior.groupby("so").agg(
        raw_sc_prior_delivery_mt=("raw_sc_vol_mt", "sum")
    ).reset_index()


def _prior_delivery_mask(df: pd.DataFrame, data_date: str) -> pd.Series:
    if not data_date:
        return pd.Series(False, index=df.index)
    start, end = _previous_month_window(data_date)
    return (
        df["loading_date_sc"].between(start, end, inclusive="both")
        & df["delivery_pcs"].notna()
        & (df["delivery_pcs"] >= 0)
    )


def _previous_month_window(data_date: str) -> tuple[pd.Timestamp, pd.Timestamp]:
    current = pd.to_datetime(data_date)
    start = (current.replace(day=1) - pd.DateOffset(months=1)).normalize()
    end = (current.replace(day=1) - pd.DateOffset(days=1)).normalize()
    return start, end


def _order_type_label(values: pd.Series) -> str:
    types = [v for v in values.dropna().unique().tolist() if v]
    if len(types) == 1:
        return types[0]
    return "Mixed"


def _audit_frame(df: pd.DataFrame) -> pd.DataFrame:
    cols = [
        "so", "plant", "cluster", "end_customer_id", "raw_sc_vol_mt",
        "adjusted_sc_vol_mt", "adjustment_factor", "order_type", "in_baseline",
        "customer_mapping_matched", "delivery_pcs", "loading_date_sc",
        "sc_prior_delivery_eligible",
    ]
    available = [c for c in cols if c in df.columns]
    return df[available].copy()


def _load_customer_keys(file_path: str) -> set:
    df = pd.read_excel(file_path, sheet_name="dim->new cc")
    df.columns = df.columns.str.strip()
    key_col = _find_column(df, "new customer code")
    return set(df[key_col].apply(_to_code_str).dropna()) - {""}


def _to_so_str(value) -> str:
    """Convert SO value (often float) to clean string."""
    if pd.isna(value):
        return ""
    try:
        return str(int(float(value)))
    except (ValueError, TypeError):
        return str(value).strip()


def _to_code_str(value) -> str:
    """Convert Excel-stored customer/material code to stable string."""
    if pd.isna(value):
        return ""
    text = str(value).strip()
    try:
        number = float(text)
        if number.is_integer():
            return str(int(number))
    except (ValueError, TypeError):
        pass
    if text.endswith(".0"):
        return text[:-2]
    return text


def _map_plant(value) -> str:
    """Map Supply From to plant code."""
    val = str(value).strip().lower()
    if "kunshan" in val or "ks" in val:
        return "KS"
    elif "indonesia" in val or "idn" in val:
        return "IDN"
    return "Unknown"


def _normalize_col_name(value) -> str:
    """Normalize headers so embedded newlines/case changes do not break lookup."""
    return " ".join(str(value).strip().lower().split())


def _find_column(df: pd.DataFrame, name: str, fallback: str = None, col_position: str = None) -> str:
    """Find column by normalized name, falling back to partial match."""
    name_lower = _normalize_col_name(name)
    for col in df.columns:
        if _normalize_col_name(col) == name_lower:
            return col
    # Partial match
    for col in df.columns:
        if name_lower in _normalize_col_name(col):
            return col
    # Try fallback
    if fallback:
        return _find_column(df, fallback)
    raise ValueError(f"Column '{name}' not found. Available: {list(df.columns)}")
