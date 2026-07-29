# compare_addresses.py
import os
import sys
import re
import argparse
from typing import Tuple
import pandas as pd
from rapidfuzz import fuzz
import unidecode
from tqdm import tqdm

# --- Configuration ---
INPUT_PATH_DEFAULT = r"\\virgo\Work\OpControl\2016\Vlad\Adhoc\Second FBB 2025_2026\Add_AI.xlsx"
SHEET_NAME = None  # None => first sheet; or set sheet name or index
ADDRESS_COL = "ADDRESS"        # adjust if your column name differs
ADDRESS_OLD_COL = "Address_Old"  # adjust if your column name differs
OUTPUT_CSV = "addresses_compared.csv"
CHUNK_ROWS = None  # set to integer to process in chunks for very large files

# Thresholds for flags
FLAG_THRESHOLDS = {
    "not_ok_similarity": 95.0,
    "review_min": 80.0
}

# Abbreviation expansions (add or edit as needed)
ABBREVIATIONS = {
    r"\bst\b": "street",
    r"\bst\.?\b": "street",
    r"\brd\b": "road",
    r"\brd\.?\b": "road",
    r"\bave\b": "avenue",
    r"\bave\.?\b": "avenue",
    r"\bblvd\b": "boulevard",
    r"\bapt\b": "apartment",
    r"\bapt\.?\b": "apartment",
    r"\bfl\b": "floor",
    r"\bno\b": "number",
    r"#": "number "
}

# --- Normalization function ---
def normalize_address(s: str) -> str:
    if pd.isna(s):
        return ""
    s = str(s)
    s = unidecode.unidecode(s)            # remove diacritics
    s = s.lower().strip()
    # replace punctuation with space except # which we handle in ABBREVIATIONS
    s = re.sub(r"[^\w\s#]", " ", s)
    # expand abbreviations
    for pat, repl in ABBREVIATIONS.items():
        s = re.sub(pat, repl, s)
    # collapse multiple spaces
    s = re.sub(r"\s+", " ", s).strip()
    return s

# --- Similarity function ---
def similarity_percent(a: str, b: str) -> float:
    if not a and not b:
        return 100.0
    # token_set_ratio handles word order and duplicates
    t = fuzz.token_set_ratio(a, b)
    l = fuzz.ratio(a, b)
    score = 0.7 * t + 0.3 * l
    return round(score, 1)

# --- Flagging function ---
def flag_row(a_norm: str, b_norm: str, sim: float) -> str:
    if a_norm == b_norm:
        return "Not Ok"
    if sim >= FLAG_THRESHOLDS["not_ok_similarity"]:
        return "Not Ok"
    if sim >= FLAG_THRESHOLDS["review_min"]:
        return "Review"
    return "Ok"

# --- Main processing ---
def process_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    # Ensure columns exist
    if ADDRESS_COL not in df.columns or ADDRESS_OLD_COL not in df.columns:
        raise KeyError(f"Expected columns '{ADDRESS_COL}' and '{ADDRESS_OLD_COL}' in the sheet. Found: {list(df.columns)}")
    # Normalize
    tqdm.pandas(desc="Normalizing addresses")
    df["A_norm"] = df[ADDRESS_COL].progress_apply(normalize_address)
    df["B_norm"] = df[ADDRESS_OLD_COL].progress_apply(normalize_address)
    # Similarity and flag
    sims = []
    flags = []
    for a, b in tqdm(zip(df["A_norm"], df["B_norm"]), total=len(df), desc="Computing similarity"):
        sim = similarity_percent(a, b)
        sims.append(sim)
        flags.append(flag_row(a, b, sim))
    df["SimilarityPercent"] = sims
    df["MatchFlag"] = flags
    # Keep order and return
    cols_out = [ADDRESS_COL, ADDRESS_OLD_COL, "MatchFlag", "SimilarityPercent", "A_norm", "B_norm"]
    return df[cols_out]

def read_and_process(input_path: str, sheet_name=None, chunk_rows=None) -> pd.DataFrame:
    # If chunk_rows is None, read whole sheet
    if chunk_rows is None:
        df = pd.read_excel(input_path, sheet_name=sheet_name, engine="openpyxl", dtype=str)
        return process_dataframe(df)
    # If chunking requested (for very large files), iterate
    reader = pd.read_excel(input_path, sheet_name=sheet_name, engine="openpyxl", dtype=str, chunksize=chunk_rows)
    out_frames = []
    for chunk in tqdm(reader, desc="Reading chunks"):
        out_frames.append(process_dataframe(chunk))
    return pd.concat(out_frames, ignore_index=True)

def main():
    parser = argparse.ArgumentParser(description="Compare Address and Address_Old columns row by row.")
    parser.add_argument("--input", "-i", default=INPUT_PATH_DEFAULT, help="Path to Excel file")
    parser.add_argument("--sheet", "-s", default=SHEET_NAME, help="Sheet name or index (optional)")
    parser.add_argument("--output", "-o", default=OUTPUT_CSV, help="Output CSV filename")
    parser.add_argument("--chunk", "-c", type=int, default=CHUNK_ROWS, help="Rows per chunk for large files (optional)")
    args = parser.parse_args()

    input_path = args.input
    if not os.path.exists(input_path):
        print(f"ERROR: Input file not found: {input_path}", file=sys.stderr)
        sys.exit(2)

    try:
        result_df = read_and_process(input_path, sheet_name=args.sheet, chunk_rows=args.chunk)
    except Exception as e:
        print("ERROR during processing:", e, file=sys.stderr)
        raise

    # Save results
    result_df.to_csv(args.output, index=False, encoding="utf-8")
    # Print summary
    total = len(result_df)
    not_ok = (result_df["MatchFlag"] == "Not Ok").sum()
    review = (result_df["MatchFlag"] == "Review").sum()
    ok = (result_df["MatchFlag"] == "Ok").sum()
    print(f"Processed {total} rows. Not Ok: {not_ok}. Review: {review}. Ok: {ok}. Output: {args.output}")

if __name__ == "__main__":
    main()
