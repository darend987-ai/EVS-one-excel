# logic/allele_compare.py

import os
from io import BytesIO
import hashlib
from datetime import datetime
from decimal import Decimal, InvalidOperation
import zipfile

import pandas as pd


# ==========================================================
# NORMALIZATION & HELPERS
# ==========================================================
def canonical_allele(val):
    """Normalize allele values to canonical strings (10.0→'10', 9.30→'9.3')."""
    if val is None:
        return None
    s = str(val).strip()
    if s == "" or s.lower() == "nan" or s == "-":
        return None
    try:
        d = Decimal(s)
    except (InvalidOperation, ValueError):
        return s
    if d == d.to_integral():
        return str(int(d))
    s_norm = format(d.normalize(), "f")
    if s_norm.endswith("."):
        s_norm = s_norm[:-1]
    return s_norm


def allele_sort_key(a):
    """Sort alleles numerically when possible."""
    try:
        d = Decimal(a)
        return (0, float(d), "")
    except Exception:
        return (1, float("inf"), str(a))


def unique_sheet_names(base_name: str, idx: int, existing_names=None):
    """
    Generate two Excel-safe sheet names WITHOUT hashes:
      PL_<shortname>  and  SUM_<shortname>
    If trimming to 31 chars causes a collision, append _<idx>.
    existing_names: iterable of sheet names already in the workbook.
    """
    invalid = "\\/?*:[]\""
    safe = "".join(ch if ch not in invalid else "_" for ch in str(base_name)).strip()

    # Leave room for "PL_" / "SUM_" prefixes so total max is 31
    short = safe[:28] if len(safe) > 28 else safe

    pl = f"PL_{short}"[:31]
    sm = f"SUM_{short}"[:31]

    existing = set(existing_names or [])

    # If either already exists, append _<idx> (and trim again to 31)
    if pl in existing:
        pl = (f"PL_{short}_{idx}")[:31]
    if sm in existing:
        sm = (f"SUM_{short}_{idx}")[:31]

    return pl, sm


# ==========================================================
# COMPARISON LOGIC
# ==========================================================
def compare_profiles(contrib_profile, suspect_profile, evidence_profile):
    """Return detailed (per-locus) and summary DataFrames, with Expected column."""
    # Loci universe (preserve previous behavior)
    all_loci = sorted(
        set(contrib_profile.keys()) |
        set(suspect_profile.keys()) |
        set(evidence_profile.keys())
    )

    rows = []
    summary_counts = {
        "Obligate_of_Suspect": 0,
        "Obligate_of_Assumed_Contributor": 0,
        "Shared_All_Three": 0,
        "Missing_from_Suspect": 0,
        "Missing_from_Assumed_Contributor": 0,
        "Foreign_Alleles": 0,
    }

    # ---- helpers
    def allele_sort_key(a):
        from decimal import Decimal
        try:
            d = Decimal(a)
            return (0, float(d), "")
        except Exception:
            return (1, float("inf"), str(a))

    def fmt(xs):
        if not xs:
            return "-"
        return ", ".join(sorted(xs, key=allele_sort_key))

    # ---- per-locus (observed) computation
    for locus in all_loci:
        c = contrib_profile.get(locus, set())
        s = suspect_profile.get(locus, set())
        e = evidence_profile.get(locus, set())

        must_from_suspect = (e & s) - c
        must_from_contrib = (e & c) - s
        shared_all_three = e & s & c
        missing_from_evidence_suspect = s - e
        missing_from_evidence_contrib = c - e
        foreign_alleles = e - (s | c)

        rows.append({
            "Locus": locus,
            "Evidence Alleles": fmt(e),
            "Suspect Alleles": fmt(s),
            "Contributor Alleles": fmt(c),
            "Obligate_of_Suspect": fmt(must_from_suspect),
            "Obligate_of_Assumed_Contributor": fmt(must_from_contrib),
            "Shared_All_Three": fmt(shared_all_three),
            "Missing_from_Suspect": fmt(missing_from_evidence_suspect),
            "Missing_from_Assumed_Contributor": fmt(missing_from_evidence_contrib),
            "Foreign_Alleles": fmt(foreign_alleles),
            "Obligate_of_Suspect_Count": len(must_from_suspect),
            "Obligate_of_Assumed_Contributor_Count": len(must_from_contrib),
            "Shared_All_Three_Count": len(shared_all_three),
            "Missing_from_Suspect_Count": len(missing_from_evidence_suspect),
            "Missing_from_Assumed_Contributor_Count": len(missing_from_evidence_contrib),
            "Foreign_Alleles_Count": len(foreign_alleles),
        })

        summary_counts["Obligate_of_Suspect"] += len(must_from_suspect)
        summary_counts["Obligate_of_Assumed_Contributor"] += len(must_from_contrib)
        summary_counts["Shared_All_Three"] += len(shared_all_three)
        summary_counts["Missing_from_Suspect"] += len(missing_from_evidence_suspect)
        summary_counts["Missing_from_Assumed_Contributor"] += len(missing_from_evidence_contrib)
        summary_counts["Foreign_Alleles"] += len(foreign_alleles)

    df_per_locus = pd.DataFrame(rows)

    # ---- Expected values (independent of evidence)
    # Totals across all loci for suspect and contributor
    suspect_total = sum(len(suspect_profile.get(l, set())) for l in all_loci)
    contrib_total = sum(len(contrib_profile.get(l, set())) for l in all_loci)
    shared_total = sum(
        len(suspect_profile.get(l, set()) & contrib_profile.get(l, set()))
        for l in all_loci
    )

    # Build summary with Observed and Expected
    categories = list(summary_counts.keys())
    observed = [summary_counts[k] for k in categories]
    expected_map = {
        "Obligate_of_Suspect": suspect_total,
        "Obligate_of_Assumed_Contributor": contrib_total,
        "Shared_All_Three": shared_total,
        # others left blank
    }
    expected = [expected_map.get(k, "") for k in categories]

    df_summary = pd.DataFrame({
        "Category": categories,
        "Observed": observed,   # previously 'Count'
        "Expected": expected
    })

    return df_per_locus, df_summary


# ==========================================================
# INPUT HANDLING (single-sheet wide format)
# ==========================================================
IGNORED_LOCI = {"amelogenin", "amel", "dys391", "qs1", "qs2"}
IGNORED_SAMPLE_KEYWORDS = ("pos", "neg", "ladder", "hidi")

SUSPECT_KEYWORDS = ("POI-ref", "POI", "suspect", "sus")
ASSUMED_KEYWORDS = ("ref", "assumed", "victim", "contributor")


def _engine_for(filename: str):
    ext = os.path.splitext(filename)[1].lower()
    if ext == ".xlsx":
        return "openpyxl"
    if ext == ".xls":
        return "xlrd"
    return None


def load_excel_sheet(uploaded_file):
    """
    Read the first sheet from an uploaded file-like object:
      - Row 1 => headers (locus names)
      - Col 1 => sample names (no header in A1)
    """
    engine = _engine_for(getattr(uploaded_file, "name", ""))
    df = pd.read_excel(uploaded_file, header=0, dtype=object, engine=engine)
    first_col = df.columns[0]
    if str(first_col).strip().lower() != "sample":
        df = df.rename(columns={first_col: "Sample"})
    return df


def filter_loci_and_samples(df):
    """Filter out ignored loci and control/ladder samples."""
    keep_cols = ["Sample"]
    for col in df.columns[1:]:
        if str(col).strip().lower() not in IGNORED_LOCI:
            keep_cols.append(col)
    df = df[keep_cols].copy()

    pat = "|".join(IGNORED_SAMPLE_KEYWORDS)
    mask_bad = df["Sample"].astype(str).str.contains(pat, case=False, na=False)
    df = df[~mask_bad].reset_index(drop=True)
    return df


def classify_rows(df):
    """
    Classify rows by sample name:
      - Suspect: any of SUSPECT_KEYWORDS
      - Assumed: any of ASSUMED_KEYWORDS but NOT suspect
      - Evidence: everything else
    If multiples exist for suspect/assumed, take the first alphabetically (silent).
    """
    names = df["Sample"].astype(str)
    suspect_mask = names.str.contains("|".join(SUSPECT_KEYWORDS), case=False, na=False)
    assumed_mask = names.str.contains("|".join(ASSUMED_KEYWORDS), case=False, na=False) & ~suspect_mask
    evidence_mask = ~(suspect_mask | assumed_mask)

    suspect_df = df[suspect_mask].sort_values("Sample")
    assumed_df = df[assumed_mask].sort_values("Sample")
    evidence_df = df[evidence_mask].sort_values("Sample")

    suspect_row = suspect_df.iloc[0] if not suspect_df.empty else None
    assumed_row = assumed_df.iloc[0] if not assumed_df.empty else None
    return suspect_row, assumed_row, evidence_df


def _split_alleles(raw):
    """Robust allele parser for comma-separated cells."""
    if raw is None or (isinstance(raw, float) and pd.isna(raw)) or (isinstance(raw, str) and raw.strip().lower() in ("", "nan", "-")):
        return []
    parts = str(raw).split(",")
    return [p.strip() for p in parts if p.strip() != ""]


def row_to_profile(row, loci_cols):
    """Convert one row into a {locus: set(alleles)} dict."""
    profile = {}
    for locus in loci_cols:
        raw = row.get(locus)
        alleles = _split_alleles(raw)
        c_alleles = []
        for a in alleles:
            ca = canonical_allele(a)
            if ca is not None:
                c_alleles.append(ca)
        profile[str(locus).strip()] = set(c_alleles)
    return profile


# ==========================================================
# SINGLE-FILE & BATCH RUNNERS (in-memory)
# ==========================================================
def run_comparison_from_file(uploaded_file):
    """
    Process a single uploaded Excel file-like object.
    Returns:
      - results: dict[evidence_name] = (df_per_locus, df_summary)
      - output_excel: BytesIO .xlsx file with two sheets per evidence
    """
    df0 = load_excel_sheet(uploaded_file)
    df = filter_loci_and_samples(df0)

    suspect_row, assumed_row, evidence_df = classify_rows(df)
    if suspect_row is None:
        raise RuntimeError("No suspect row found (contains one of: POI-ref, POI, suspect, sus).")
    if assumed_row is None:
        raise RuntimeError("No assumed contributor row found (contains one of: ref, assumed, victim, contributor).")
    if evidence_df.empty:
        raise RuntimeError("No evidence rows found (names without suspect/assumed keywords).")

    loci_cols = [c for c in df.columns if c != "Sample"]
    suspect_profile = row_to_profile(suspect_row, loci_cols)
    assumed_profile = row_to_profile(assumed_row, loci_cols)

    # Write an in-memory Excel with two sheets per evidence
    output = BytesIO()
    results = {}
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
    # Track existing names locally to avoid re-reading every time
        existing = set(getattr(writer.book, "sheetnames", []))

        for idx, (_, ev_row) in enumerate(evidence_df.iterrows(), start=1):
            ev_name = str(ev_row["Sample"]).strip()
            ev_profile = row_to_profile(ev_row, loci_cols)

            df_per_locus, df_summary = compare_profiles(
                assumed_profile,  # assumed contributor
                suspect_profile,  # suspect
                ev_profile        # evidence
            )

            results[ev_name] = (df_per_locus, df_summary)

        # ✅ Clean sheet names (no hashes), add _<idx> only if needed
            per_locus_sheet, summary_sheet = unique_sheet_names(ev_name, idx, existing_names=existing)

        # Write both sheets
            df_per_locus.to_excel(writer, sheet_name=per_locus_sheet, index=False)
            df_summary.to_excel(writer, sheet_name=summary_sheet, index=False)

        # Update local set so subsequent names consider these
            existing.add(per_locus_sheet)
            existing.add(summary_sheet)

    output.seek(0)
    return results, output


def run_batch_comparison(uploaded_files):
    """
    Process multiple uploaded Excel files and return a timestamped ZIP (BytesIO).
    Each result Excel inside the ZIP is named:
      <original_filename_without_ext>_comparison_results.xlsx
    """
    zip_bytes = BytesIO()
    with zipfile.ZipFile(zip_bytes, "w", compression=zipfile.ZIP_DEFLATED) as zipf:
        for uf in uploaded_files:
            # Run comparison for this file
            _, excel_bytes = run_comparison_from_file(uf)
            base = os.path.splitext(os.path.basename(uf.name))[0]
            out_name = f"{base}_comparison_results.xlsx"
            zipf.writestr(out_name, excel_bytes.getvalue())
    zip_bytes.seek(0)
    # Return both ZIP bytes and number of files processed for UI messaging
    return zip_bytes, len(uploaded_files)


def timestamp_str():
    return datetime.now().strftime("%Y-%m-%d_%H%M%S")
