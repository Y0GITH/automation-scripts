"""
Excel/CSV Data Profiler & Column Similarity Finder
------------------------------------------------
Scans a folder (and every subfolder inside it) for data files
(.xlsx, .xls, .xlsm, .csv, .tsv, .txt), treats every sheet/file as
one "table", auto-generates a data dictionary, finds similar/related
columns across tables (for star-schema modeling), and runs outlier +
clustering + correlation analysis on numeric columns.

It only READS the source files you point it at, and WRITES a brand-new
output file (analysis.xlsx) into a folder you choose.

Requirements:
    pip install pandas numpy rapidfuzz scikit-learn openpyxl
"""

import itertools
import os
import re
from pathlib import Path

import numpy as np
import pandas as pd
from rapidfuzz import fuzz
from sklearn.cluster import KMeans
from sklearn.metrics import silhouette_score

# =========================== CONFIG ===========================
SAMPLE_SIZE = 500  # rows sampled per column for similarity checks (kept small = fast)
SIMILARITY_THRESHOLD = 30  # only keep column pairs scoring >= this % match
NAME_WEIGHT = 0.4  # weight given to column-name similarity
VALUE_WEIGHT = 0.6  # weight given to value/pattern similarity
MAX_CLUSTERS = 6  # upper bound tried when auto-picking cluster count
CLUSTER_SAMPLE_SIZE = 20000  # cap on rows used for KMeans/silhouette on big columns
TEXT_ENCODINGS = ("utf-8", "cp1252", "latin-1")  # tried in order for csv/tsv/txt files

DATA_EXTENSIONS = (".xlsx", ".xls", ".xlsm", ".csv", ".tsv", ".txt")
OUTPUT_FILE_NAME = "analysis.xlsx"
DEFAULT_OUTPUT_FOLDER_NAME = "quick tables summery"
# ================================================================


# ===================== 0. FILE DISCOVERY & LOADING =====================


def discover_files(root_folder, extensions=DATA_EXTENSIONS):
    """Recursively find every data file under root_folder (all subfolders included)."""
    root = Path(root_folder)
    found = []
    for path in root.rglob("*"):
        if not path.is_file():
            continue
        if path.name.startswith("~$"):  # skip Excel lock files
            continue
        if path.suffix.lower() in extensions:
            found.append(path)
    return sorted(found)


def make_unique_name(base_name, existing_names):
    """Avoid clobbering two tables that would otherwise get the same name."""
    name = base_name
    counter = 2
    while name in existing_names:
        name = f"{base_name} ({counter})"
        counter += 1
    return name


def read_delimited_text(path):
    """Try to read a .csv/.tsv/.txt file, sniffing the delimiter and encoding."""
    suffix = path.suffix.lower()
    last_error = None
    for encoding in TEXT_ENCODINGS:
        try:
            if suffix == ".tsv":
                return pd.read_csv(path, sep="\t", encoding=encoding)
            if suffix == ".csv":
                return pd.read_csv(path, encoding=encoding)
            # .txt: let pandas try to sniff the separator
            try:
                return pd.read_csv(path, sep=None, engine="python", encoding=encoding)
            except (UnicodeDecodeError, pd.errors.ParserError):
                return pd.read_csv(path, sep="\t", encoding=encoding)
        except UnicodeDecodeError as e:
            last_error = e
            continue  # try the next encoding
    raise last_error


def load_all_tables(root_folder):
    """
    Walk root_folder (and every subfolder), load every data file found,
    and return a dict of {table_name: DataFrame} - exactly the shape the
    rest of this script already expects (each sheet/file = one table).
    """
    files = discover_files(root_folder)
    print(f"Found {len(files)} data file(s) under: {root_folder}")

    sheets = {}
    for f in files:
        try:
            if f.suffix.lower() in (".xlsx", ".xls", ".xlsm"):
                workbook = pd.read_excel(f, sheet_name=None)
                for sheet_name, df in workbook.items():
                    label = f.stem if len(workbook) == 1 else f"{f.stem}__{sheet_name}"
                    label = make_unique_name(label, sheets)
                    sheets[label] = df
                    print(
                        f"  Loaded: {f.name} -> sheet '{sheet_name}' ({len(df)} rows)"
                    )
            else:
                df = read_delimited_text(f)
                label = make_unique_name(f.stem, sheets)
                sheets[label] = df
                print(f"  Loaded: {f.name} ({len(df)} rows)")
        except Exception as e:
            print(f"  Skipped {f} due to read error: {e}")

    return sheets


# ===================== INTERACTIVE PROMPTS =====================


def clean_path_input(raw):
    """Strip stray quotes/whitespace people paste in from Windows Explorer."""
    return raw.strip().strip('"').strip("'")


def prompt_input_folder():
    while True:
        raw = input(
            "\nFolder to scan for data files (CSV/Excel/text, includes subfolders): "
        )
        path = clean_path_input(raw)
        if os.path.isdir(path):
            return path
        print(f"  '{path}' is not a valid folder. Please try again.")


def prompt_yes_no(question, default_yes=True):
    suffix = "Y/n" if default_yes else "y/N"
    while True:
        raw = input(f"{question} ({suffix}): ").strip().lower()
        if raw == "":
            return default_yes
        if raw in ("y", "yes"):
            return True
        if raw in ("n", "no"):
            return False
        print("  Please answer y or n.")


def prompt_output_location():
    while True:
        raw = input(
            "\nWhere should the output folder be created? (e.g. C:\\Users\\Lenovo\\Downloads): "
        )
        base = clean_path_input(raw)
        if os.path.isdir(base):
            break
        create_anyway = (
            input(f"  '{base}' doesn't exist yet. Create it too? (y/n): ")
            .strip()
            .lower()
        )
        if create_anyway == "y":
            break
        print("  Let's try again.")

    folder_name = input(
        f"Name for the new output folder [{DEFAULT_OUTPUT_FOLDER_NAME}]: "
    ).strip()
    if not folder_name:
        folder_name = DEFAULT_OUTPUT_FOLDER_NAME
    output_dir = Path(base) / folder_name
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / OUTPUT_FILE_NAME
    print(f"  Report will be saved to: {output_path}")
    return output_path


# ===================== 1. DATA DICTIONARY =====================


def infer_dtype_label(series):
    """Return a friendly data-type label for a column."""
    if pd.api.types.is_bool_dtype(series):
        return "boolean"
    if pd.api.types.is_integer_dtype(series):
        return "integer"
    if pd.api.types.is_float_dtype(series):
        return "decimal"
    if pd.api.types.is_datetime64_any_dtype(series):
        return "date/time"
    non_null = series.dropna()
    if len(non_null) == 0:
        return "unknown (all blank)"
    numeric_try = pd.to_numeric(non_null, errors="coerce")
    if numeric_try.notna().mean() > 0.9:
        return "numeric (stored as text)"
    date_try = pd.to_datetime(non_null, errors="coerce")
    if date_try.notna().mean() > 0.9:
        return "date (stored as text)"
    return "text"


def count_errors(series, dtype_label):
    """Count values that don't match the column's dominant inferred type."""
    non_null = series.dropna()
    if len(non_null) == 0:
        return 0
    if dtype_label == "numeric (stored as text)":
        return int(pd.to_numeric(non_null, errors="coerce").isna().sum())
    if dtype_label == "date (stored as text)":
        return int(pd.to_datetime(non_null, errors="coerce").isna().sum())
    return 0


def count_blank_strings(series):
    """Count empty-string / whitespace-only entries (distinct from real NaN)."""
    if series.dtype != object:
        return 0
    return int(series.astype(str).str.strip().eq("").sum())


def build_data_dictionary(sheets):
    rows = []
    slno = 1
    for t_no, (table_name, df) in enumerate(sheets.items(), start=1):
        for c_no, col in enumerate(df.columns, start=1):
            series = df[col]
            dtype_label = infer_dtype_label(series)
            rows.append(
                {
                    "slno": slno,
                    "table no": t_no,
                    "table name": table_name,
                    "column no": c_no,
                    "column name": col,
                    "data type": dtype_label,
                    "total row": len(series),
                    "is na": int(series.isna().sum()),
                    "is error": count_errors(series, dtype_label),
                    "is null": count_blank_strings(series),
                }
            )
            slno += 1
    return pd.DataFrame(rows)


# ================ 2. DIMENSION / CARDINALITY / PRIMARY KEY ================


def build_dimension_and_key_report(sheets):
    rows = []
    for table_name, df in sheets.items():
        dup_count = int(df.duplicated().sum())
        for col in df.columns:
            series = df[col]
            nunique = series.nunique(dropna=True)
            total = len(series)
            is_pk_candidate = (
                (series.isna().sum() == 0) and (nunique == total) and total > 0
            )
            rows.append(
                {
                    "table name": table_name,
                    "column name": col,
                    "distinct values (cardinality)": nunique,
                    "duplicate rows in table": dup_count,
                    "primary key candidate": "YES" if is_pk_candidate else "",
                }
            )
    return pd.DataFrame(rows)


# ===================== 3. COLUMN SIMILARITY =====================


def sample_values(series, n=SAMPLE_SIZE):
    """Return a set of up to n sampled, stringified, non-null unique values."""
    non_null = series.dropna()
    if len(non_null) > n:
        non_null = non_null.sample(n=n, random_state=42)
    return set(non_null.astype(str).str.strip().str.lower())


def value_pattern(value):
    """
    Turn a value into a shape signature: runs of letters become L<len>,
    runs of digits become D<len>, everything else is kept as-is.
    e.g. "abc123" -> "L3D3", "xyz456" -> "L3D3" (same shape, different codes),
    "AB-9981" -> "L2-D4".
    """
    s = str(value)
    s = re.sub(r"[A-Za-z]+", lambda m: f"L{len(m.group(0))}", s)
    s = re.sub(r"\d+", lambda m: f"D{len(m.group(0))}", s)
    return s


def pattern_set(value_sample):
    """Convert a sample of raw values into the set of shape signatures they produce."""
    return {value_pattern(v) for v in value_sample}


def value_similarity(set_a, set_b):
    """Jaccard similarity (%) between two sets (raw values OR pattern shapes)."""
    if not set_a or not set_b:
        return 0.0
    intersection = len(set_a & set_b)
    union = len(set_a | set_b)
    return (intersection / union) * 100 if union else 0.0


def build_similarity_report(sheets):
    col_samples = {}
    col_patterns = {}
    for table_name, df in sheets.items():
        for col in df.columns:
            raw_sample = sample_values(df[col])
            col_samples[(table_name, col)] = raw_sample
            col_patterns[(table_name, col)] = pattern_set(raw_sample)

    results = []
    keys = list(col_samples.keys())
    for (t1, c1), (t2, c2) in itertools.combinations(keys, 2):
        if t1 == t2:
            continue  # only compare across DIFFERENT tables (relationship discovery)
        name_sim = fuzz.token_sort_ratio(str(c1), str(c2))
        val_sim = value_similarity(col_samples[(t1, c1)], col_samples[(t2, c2)])
        pattern_sim = value_similarity(col_patterns[(t1, c1)], col_patterns[(t2, c2)])
        # Use whichever signal is stronger: literal overlap catches real shared
        # keys (e.g. matching customer IDs); pattern overlap catches columns
        # that are clearly "the same kind of code" even when the actual values differ.
        best_value_signal = max(val_sim, pattern_sim)
        combined = (NAME_WEIGHT * name_sim) + (VALUE_WEIGHT * best_value_signal)
        if combined >= SIMILARITY_THRESHOLD:
            if combined >= 70:
                if val_sim >= 50:
                    verdict = "possible key / FK match"
                else:
                    verdict = "same format/pattern, values differ - verify"
            elif combined >= 50:
                verdict = "worth reviewing"
            else:
                verdict = "weak match"
            results.append(
                {
                    "table 1": t1,
                    "column 1": c1,
                    "table 2": t2,
                    "column 2": c2,
                    "name similarity %": round(name_sim, 1),
                    "value similarity %": round(val_sim, 1),
                    "pattern similarity %": round(pattern_sim, 1),
                    "match %": round(combined, 1),
                    "likely relationship": verdict,
                }
            )
    result_df = pd.DataFrame(results)
    if not result_df.empty:
        result_df.sort_values("match %", ascending=False, inplace=True)
        result_df.reset_index(drop=True, inplace=True)
    return result_df


# ===================== 4. OUTLIERS (IQR method) =====================


def build_outlier_report(sheets):
    rows = []
    for table_name, df in sheets.items():
        numeric_cols = df.select_dtypes(include=[np.number]).columns
        for col in numeric_cols:
            series = df[col].dropna()
            if len(series) < 4:
                continue
            q1, q3 = series.quantile([0.25, 0.75])
            iqr = q3 - q1
            lower = q1 - 1.5 * iqr
            upper = q3 + 1.5 * iqr
            outliers = series[(series < lower) | (series > upper)]
            rows.append(
                {
                    "table name": table_name,
                    "column name": col,
                    "normal range (low)": round(lower, 2),
                    "normal range (high)": round(upper, 2),
                    "outlier count": len(outliers),
                    "outlier %": round(len(outliers) / len(series) * 100, 2),
                }
            )
    return pd.DataFrame(rows)


# ===================== 5. CLUSTERING (per numeric column) =====================


def best_k_for_column(values, max_k=MAX_CLUSTERS):
    """Pick the cluster count with the best silhouette score."""
    best_k, best_score = 2, -1
    n = len(values)
    for k in range(2, min(max_k, n - 1) + 1):
        km = KMeans(n_clusters=k, n_init=10, random_state=42)
        labels = km.fit_predict(values)
        if len(set(labels)) < 2:
            continue
        score = silhouette_score(values, labels)
        if score > best_score:
            best_k, best_score = k, score
    return best_k


def build_cluster_report(sheets):
    rows = []
    for table_name, df in sheets.items():
        numeric_cols = df.select_dtypes(include=[np.number]).columns
        for col in numeric_cols:
            series = df[col].dropna()
            if len(series) < 10 or series.nunique() < 3:
                continue
            # Cap rows used for KMeans/silhouette - both get very slow on
            # multi-million-row columns and a sample is plenty for this purpose.
            fit_series = (
                series.sample(n=CLUSTER_SAMPLE_SIZE, random_state=42)
                if len(series) > CLUSTER_SAMPLE_SIZE
                else series
            )
            values = fit_series.values.reshape(-1, 1)
            k = best_k_for_column(values)
            km = KMeans(n_clusters=k, n_init=10, random_state=42)
            km.fit(values)
            # Assign ALL rows (not just the sample) to their nearest cluster
            # so the reported ranges reflect the full column.
            all_labels = km.predict(series.values.reshape(-1, 1))
            for cluster_id in range(k):
                cluster_vals = series[all_labels == cluster_id]
                if len(cluster_vals) == 0:
                    continue
                rows.append(
                    {
                        "table name": table_name,
                        "column name": col,
                        "clusters chosen": k,
                        "cluster id": cluster_id,
                        "cluster size": len(cluster_vals),
                        "range low": round(cluster_vals.min(), 2),
                        "range high": round(cluster_vals.max(), 2),
                    }
                )
    return pd.DataFrame(rows)


# ===================== 6. CORRELATION (numeric columns, per table) =====================


def build_correlation_report(sheets):
    rows = []
    for table_name, df in sheets.items():
        numeric_df = df.select_dtypes(include=[np.number])
        if numeric_df.shape[1] < 2:
            continue
        corr = numeric_df.corr()
        for c1, c2 in itertools.combinations(corr.columns, 2):
            rows.append(
                {
                    "table name": table_name,
                    "column 1": c1,
                    "column 2": c2,
                    "correlation": round(corr.loc[c1, c2], 3),
                }
            )
    return pd.DataFrame(rows)


# ===================== WRITE OUTPUT =====================


def write_report(
    output_path,
    dictionary_df,
    dim_key_df,
    similarity_df,
    outlier_df,
    cluster_df,
    corr_df,
):
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        dictionary_df.to_excel(writer, sheet_name="Data Dictionary", index=False)
        dim_key_df.to_excel(writer, sheet_name="Cardinality & Keys", index=False)
        similarity_df.to_excel(writer, sheet_name="Column Similarity", index=False)
        outlier_df.to_excel(writer, sheet_name="Outliers", index=False)
        cluster_df.to_excel(writer, sheet_name="Clusters", index=False)
        corr_df.to_excel(writer, sheet_name="Correlation", index=False)


def main():
    input_folder = prompt_input_folder()
    output_path = prompt_output_location()

    print(
        "\nClustering can take a long time on very large numeric columns "
        "(millions of rows) - similarity is always run, but sampled small "
        f"({SAMPLE_SIZE} rows/column) so it stays fast."
    )
    run_clustering = prompt_yes_no(
        "  Run clustering on numeric columns?", default_yes=True
    )

    sheets = load_all_tables(input_folder)
    if not sheets:
        print("\nNo readable data files were found - nothing to report on.")
        return

    print("\nBuilding data dictionary...")
    dictionary_df = build_data_dictionary(sheets)

    print("Checking cardinality / primary-key candidates / duplicates...")
    dim_key_df = build_dimension_and_key_report(sheets)

    print(
        f"Comparing columns across tables (sampling up to {SAMPLE_SIZE} rows each, "
        "checking both literal values and value-shape patterns)..."
    )
    similarity_df = build_similarity_report(sheets)

    print("Detecting outliers in numeric columns...")
    outlier_df = build_outlier_report(sheets)

    if run_clustering:
        print("Clustering numeric columns...")
        cluster_df = build_cluster_report(sheets)
    else:
        print("Skipping clustering (as requested).")
        cluster_df = pd.DataFrame([{"note": "Skipped by user request"}])

    print("Computing correlations...")
    corr_df = build_correlation_report(sheets)

    print(f"\nWriting report to: {output_path}")
    write_report(
        output_path,
        dictionary_df,
        dim_key_df,
        similarity_df,
        outlier_df,
        cluster_df,
        corr_df,
    )
    print("Done!")


if __name__ == "__main__":
    main()
