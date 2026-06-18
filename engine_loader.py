"""
engine_loader.py — Multi-Month Time Intelligence & Cohort Loading Pipeline (v4.0)
Parses dynamic dates, aligns tickets with delivered cohorts, and loads products.
Fixed: Integrates full order status checks and Pre/Post Delivery categories.
"""
import io
import pandas as pd
import numpy as np
import streamlit as st

# Strictly import both brand normalizer and resolved registry from engine_normalize [D]
from engine_normalize import normalize_brand_name, ProductRegistry
from engine_redistribute import (
    compute_brand_weights, redistribute_tickets,
    redistribute_subcat, build_redistribution_summary
)


def _detect_col(df, keywords, fallback=0):
    """Detects column names by keywords with absolute safe out-of-bounds index safeguards [D]."""
    cols_lower = {str(c).lower().strip(): c for c in df.columns}
    for kw in keywords:
        for col_l, col in cols_lower.items():
            if kw.lower() in col_l:
                return col
                
    if len(df.columns) == 0:
        raise ValueError("The uploaded Excel file has no columns [D].")
    if fallback >= len(df.columns):
        return df.columns[-1]
        
    return df.columns[fallback]


def parse_date_hierarchy(df, col_name, prefix):
    """Generates Date, Week, Month, Quarter, and Year columns dynamically with Epoch safeguards."""
    dt_series = pd.to_datetime(df[col_name], errors="coerce")
    
    # Bug #1 Fix: Coerce any invalid epoch date (under year 1975) to NaT [D]
    dt_series = dt_series.apply(lambda x: pd.NaT if pd.notna(x) and x.year < 1975 else x)
    
    # Safe rendering formatting [D]
    df[f"{prefix} Date"] = dt_series.dt.date
    df[f"{prefix} Date"] = df[f"{prefix} Date"].fillna("Unknown Date")
    
    df[f"{prefix} Year"] = dt_series.apply(lambda x: int(x.year) if pd.notna(x) else "Unknown Year")
    df[f"{prefix} Quarter"] = dt_series.apply(lambda x: f"{x.year}-Q{x.quarter}" if pd.notna(x) else "Unknown Quarter")
    df[f"{prefix} Month"] = dt_series.apply(lambda x: x.strftime("%B %Y") if pd.notna(x) else "Unknown Month")
    
    # Far-future default to ensure unknown sorting ranks last chronologically
    df[f"{prefix} Month Sort"] = dt_series.dt.to_period("M").fillna(pd.Period("2099-12", "M"))
    
    df[f"{prefix} Week"] = dt_series.apply(
        lambda d: f"{d.strftime('%b %Y')} Wk{min((d.day - 1) // 7 + 1, 4)}" if pd.notna(d) else "Unknown Week"
    )
    return df


def normalize_ticket_category(val):
    """Strictly normalizes raw category variants into clean PRE_DELIVERY or POST_DELIVERY keys [D]."""
    if not isinstance(val, str):
        return "POST_DELIVERY"
    s = val.strip().upper().replace("-", " ").replace("_", " ")
    if "PRE" in s:
        return "PRE_DELIVERY"
    if "POST" in s:
        return "POST_DELIVERY"
    return "POST_DELIVERY"


def load_delivered(file_bytes):
    df = pd.read_excel(io.BytesIO(file_bytes))
    df.columns = [str(c).strip() for c in df.columns]
    
    date_col = _detect_col(df, ["order_delivered_at", "delivered_at", "date"], 0)
    brand_col = _detect_col(df, ["company", "brand", "seller"], 3)
    prod_col = _detect_col(df, ["product"], 4)
    order_col = _detect_col(df, ["order_id", "orderid", "order id"], 1)
    
    # Detect delivery status column dynamically [D]
    status_col = _detect_col(df, ["status", "order_status", "delivery_status", "state"], None)
    
    out = pd.DataFrame({
        "order_id": df[order_col].astype(str).str.strip(),
        "raw_date": df[date_col],
        "raw_brand": df[brand_col].astype(str).str.strip().str.strip('"'),
        "raw_product": df[prod_col].astype(str).str.strip().str.strip('"'),
    })
    
    if status_col:
        out["is_delivered"] = df[status_col].astype(str).str.strip().str.lower() == "delivered"
    else:
        out["is_delivered"] = True # Backward compatibility fallback [D]
        
    out = parse_date_hierarchy(out, "raw_date", "Delivery")
    return out


def load_tickets(file_bytes):
    df = pd.read_excel(io.BytesIO(file_bytes))
    df.columns = [str(c).strip() for c in df.columns]
    
    date_col = _detect_col(df, ["created_at", "createdatdate", "date"], 0)
    brand_col = _detect_col(df, ["company", "brand"], 4)
    prod_col = _detect_col(df, ["product"], 3)
    order_col = _detect_col(df, ["order_id", "orderid", "order id"], 1)
    subcat_col = _detect_col(df, ["sub-category", "subcategory", "sub_cat", "sub cat", "ticket sub"], 6)
    
    # Detect ticket category column dynamically [D]
    cat_col = _detect_col(df, ["category", "ticket_category", "ticket_class", "type"], None)
    
    out = pd.DataFrame({
        "order_id": df[order_col].astype(str).str.strip(),
        "raw_date": df[date_col],
        "raw_brand": df[brand_col].astype(str).str.strip().str.strip('"'),
        "raw_product": df[prod_col].astype(str).str.strip().str.strip('"'),
        "raw_subcat":  df[subcat_col].astype(str).str.strip(),
    })
    
    # Apply category normalization to eliminate space/underscore mismatches [D]
    if cat_col:
        out["raw_category"] = df[cat_col].fillna("NULL").astype(str).str.strip()
        out["ticket_category"] = out["raw_category"].apply(normalize_ticket_category)
    else:
        out["raw_category"] = "NULL"
        out["ticket_category"] = "POST_DELIVERY" # Fallback [D]
        
    out = parse_date_hierarchy(out, "raw_date", "Ticket")
    return out


def process_pipeline(del_bytes, tick_bytes, rng_seed=42):
    rng = np.random.default_rng(rng_seed)
    prog = st.progress(0)
    status = st.empty()

    def update(pct, msg):
        prog.progress(pct)
        status.info(f"⚙️ {msg}")

    # ── Step 1: Load ──
    update(5, "Step 1 / 6 — Reading active files...")
    del_raw = load_delivered(del_bytes)
    tick_raw = load_tickets(tick_bytes)
    ORIGINAL_TICKET_COUNT = len(tick_raw)

    # ── Step 2: Normalize Brand Names ──
    update(15, "Step 2 / 6 — Normalizing brand names globally...")
    unique_del_brands = del_raw["raw_brand"].unique()
    unique_tick_brands = tick_raw["raw_brand"].unique()
    all_unique_brands = set(unique_del_brands) | set(unique_tick_brands)

    brand_map = {b: normalize_brand_name(b) for b in all_unique_brands}

    del_raw["brand"] = del_raw["raw_brand"].map(brand_map).astype(str)
    tick_raw["brand"] = tick_raw["raw_brand"].map(brand_map).astype(str)
    tick_raw["_redistributed"] = False

    del_clean = del_raw[del_raw["brand"] != "Unmapped Brand"].copy().reset_index(drop=True)

    # ── Step 3: Match Tickets with Delivery Cohorts (Bug #2 & #3 Dual Safeguards) ──
    update(35, "Step 3 / 6 — Aligning delivery cohorts with strict zero-duplication...")
    
    # Bug #2 & #3 Fix: Exclude single digit/short numerical placeholders to prevent massive duplication [D]
    valid_order_mask = (
        del_clean["order_id"].notna() & 
        (del_clean["order_id"].astype(str).str.strip() != "") & 
        (del_clean["order_id"].astype(str).str.lower() != "nan") & 
        (del_clean["order_id"].astype(str).str.len() > 3)
    )
    
    # Drop duplicates on Delivered Orders unique identifier to guarantee zero ticket duplication
    del_lookup = del_clean[valid_order_mask].drop_duplicates(subset=["order_id"]).set_index("order_id")[
        ["Delivery Date", "Delivery Week", "Delivery Month", "Delivery Quarter", "Delivery Year", "Delivery Month Sort"]
    ]
    
    # Align ticket with order delivered cohort
    tick_raw = tick_raw.join(del_lookup, on="order_id", how="left")
    
    # Fallback if order matching is absent in Delivered Orders
    tick_raw["Delivery Date"] = tick_raw["Delivery Date"].fillna(tick_raw["Ticket Date"])
    tick_raw["Delivery Week"] = tick_raw["Delivery Week"].fillna(tick_raw["Ticket Week"])
    tick_raw["Delivery Month"] = tick_raw["Delivery Month"].fillna(tick_raw["Ticket Month"])
    tick_raw["Delivery Quarter"] = tick_raw["Delivery Quarter"].fillna(tick_raw["Ticket Quarter"])
    tick_raw["Delivery Year"] = tick_raw["Delivery Year"].fillna(tick_raw["Ticket Year"])
    tick_raw["Delivery Month Sort"] = tick_raw["Delivery Month Sort"].fillna(tick_raw["Ticket Month Sort"])

    valid_mask = tick_raw["brand"] != "Unmapped Brand"
    valid_ticks = tick_raw[valid_mask].copy()

    # ── Step 4: Resolve Global Product Registry ──
    update(55, "Step 4 / 6 — Executing global resolution matching...")
    registry = ProductRegistry()
    for _, row in del_clean.iterrows():
        registry.record_delivered(row["brand"], row["raw_product"])
    for _, row in valid_ticks.iterrows():
        registry.record_ticket(row["brand"], row["raw_product"])
        
    registry.resolve()

    # Map back canonical products safely using fast dict hashes [D]
    del_clean["canonical_product"] = del_clean.apply(
        lambda r: registry.resolved_map.get(str(r["brand"]), {}).get(
            str(r["raw_product"]).strip().strip('"').strip("'"), r["raw_product"]
        ), axis=1
    )
    
    tick_raw["canonical_product"] = "Unmapped Product"
    valid_ticks["canonical_product"] = valid_ticks.apply(
        lambda r: registry.resolved_map.get(str(r["brand"]), {}).get(
            str(r["raw_product"]).strip().strip('"').strip("'"), r["raw_product"]
        ), axis=1
    )
    tick_raw.loc[valid_mask, "canonical_product"] = valid_ticks["canonical_product"].values

    brand_unmapped = tick_raw[~valid_mask].copy()

    # ── Step 5: Redistribution ──
    update(70, "Step 5 / 6 — Apportioning ticket redistributions...")
    from engine_analytics import compute_brand_summary as _bs
    base_brand_sum = _bs(del_clean, valid_ticks)
    brand_weights = compute_brand_weights(base_brand_sum, valid_ticks)

    dist_brand = redistribute_tickets(brand_unmapped, brand_weights, rng)
    if len(dist_brand) > 0:
        dist_brand["canonical_product"] = dist_brand.apply(
            lambda r: registry.resolved_map.get(str(r["brand"]), {}).get(
                str(r["raw_product"]).strip().strip('"').strip("'"), r["raw_product"]
            ), axis=1
        )

    # Merge — ALL tickets preserved
    all_ticks = pd.concat([valid_ticks, dist_brand], ignore_index=True)
    val_ok = len(all_ticks) == ORIGINAL_TICKET_COUNT

    # ── Step 6: Subcategory Normalization ──
    update(85, "Step 6 / 6 — Normalizing subcategories...")
    n_nf = int((all_ticks["raw_subcat"] == "Not Found").sum())
    n_nd = int((all_ticks["raw_subcat"] == "Need Details").sum())
    
    # Pass ticket category to ensure correct domain redistribution [D]
    all_ticks["subcat_final"] = [
        redistribute_subcat(row["raw_subcat"], row["brand"], row["canonical_product"], row["ticket_category"], rng)
        for _, row in all_ticks.iterrows()
    ]

    # Map the final ticket counts to our registry groups for correct registry reporting (Step 6 Output)
    tick_counts = all_ticks[all_ticks["brand"] != "Unmapped Brand"].groupby(["brand", "canonical_product"]).size().to_dict()
    for brand, groups in registry.final_groups.items():
        brand_str = str(brand)
        for cname in groups.keys():
            registry.final_groups[brand_str][cname]["tickets"] = tick_counts.get((brand_str, cname), 0)

    redist_summary = build_redistribution_summary(
        n_brand_nf=len(brand_unmapped),
        n_subcat_nf=n_nf,
        n_need_details=n_nd,
        brand_weights=brand_weights,
    )
    
    # Calculate date validation quality metric counts [D]
    invalid_del_dates = int(del_raw["Delivery Date"].apply(lambda x: x == "Unknown Date").sum())
    invalid_tick_dates = int(tick_raw["Ticket Date"].apply(lambda x: x == "Unknown Date").sum())

    # Extraction of ticket category value counts before and after normalization [D]
    raw_cat_counts = tick_raw["raw_category"].value_counts().to_dict()
    norm_cat_counts = tick_raw["ticket_category"].value_counts().to_dict()

    update(95, "Finalizing analytics...")
    prog.progress(100)
    status.success(f"✅ Pipeline completed — Loaded {ORIGINAL_TICKET_COUNT:,} tickets [D].")

    return {
        "del_df": del_clean,
        "tick_df": all_ticks,
        "registry": registry,
        "brand_weights": brand_weights,
        "redist_summary": redist_summary,
        "original_ticket_count": ORIGINAL_TICKET_COUNT,
        "n_unmapped_brand": len(brand_unmapped),
        "n_not_found_subcat": n_nf,
        "n_need_details": n_nd,
        "final_ticket_count": len(all_ticks),
        "validation_ok": val_ok,
        "invalid_del_dates": invalid_del_dates,
        "invalid_tick_dates": invalid_tick_dates,
        "raw_cat_counts": raw_cat_counts,
        "norm_cat_counts": norm_cat_counts
    }