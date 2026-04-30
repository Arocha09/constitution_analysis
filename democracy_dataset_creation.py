"""
democracy_merge.py
==================
Merges V-Dem and Freedom House democracy indicators.

USAGE
-----
1. Place this script in the same folder as your data files.
2. Edit the CONFIG section below to match your file names.
3. Run: python democracy_merge.py
"""

import pandas as pd
import os

# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────

# V-Dem CSV
VDEM_FILE = "data/vdem/vdem.csv"

# Freedom House Excel
FH_FILE = "data/freedom_house/All_data_FIW_2013-2024.xlsx"

# Output
OUTPUT_FILE = "merged_democracy_data.csv"

# Match year
MATCH_YEAR = 2024

# Freedom House year
FH_YEAR = 2024


# ─────────────────────────────────────────────
# V-DEM VARIABLES
# ─────────────────────────────────────────────

VDEM_KEEP = {
    "country_name": "country_name",
    "country_text_id": "country_text_id",
    "year": "year",

    "v2x_regime": "regime_type",
    "v2x_polyarchy": "electoral_democracy",
    "v2x_libdem": "liberal_democracy",
    "v2x_partipdem": "participatory_democracy",
    "v2x_delibdem": "deliberative_democracy",
    "v2x_egaldem": "egalitarian_democracy",

    "v2x_freexp_altinf": "freedom_expression",
    "v2x_frassoc_thick": "freedom_association",
    "v2xcl_rol": "rule_of_law",
    "v2x_jucon": "judicial_constraints",
    "v2xlg_legcon": "legislative_constraints",
    "v2x_civlib": "civil_liberties",
    "v2xel_frefair": "free_fair_elections",
    "v2x_suffr": "suffrage",

    "v2x_accountability": "accountability",
    "v2xnp_pres": "presidentialism",

    "v2x_gender": "gender_equality_index",
    "v2xpe_exlpol": "political_exclusion",

    "v2x_corr": "political_corruption",
    "v2x_pubcorr": "public_sector_corruption",
}

VDEM_COLS = list(VDEM_KEEP.keys())


# ─────────────────────────────────────────────
# COUNTRY NAME ALIASES
# ─────────────────────────────────────────────

COUNTRY_ALIASES = {
    "United States": "United States of America",
    "North Korea": "Korea, North",
    "Congo": "Republic of the Congo",
    "DR Congo": "Democratic Republic of the Congo",
    "Ivory Coast": "Cote d'Ivoire",
    "Eswatini": "Swaziland",
}


# ─────────────────────────────────────────────
# LOAD V-DEM
# ─────────────────────────────────────────────

def load_vdem(filepath):
    print(f"\nLoading V-Dem: {filepath}")

    available_cols = pd.read_csv(filepath, nrows=0).columns.tolist()
    cols_to_load = [c for c in VDEM_COLS if c in available_cols]

    df = pd.read_csv(filepath, usecols=cols_to_load, low_memory=False)
    df = df.rename(columns=VDEM_KEEP)

    # filter to desired year
    df = df[df["year"] == MATCH_YEAR].copy()

    # normalize country names
    df["country_name"] = (
        df["country_name"]
        .str.strip()
        .replace(COUNTRY_ALIASES)
    )

    print(f"Loaded {len(df)} V-Dem rows")
    return df


# ─────────────────────────────────────────────
# LOAD FREEDOM HOUSE
# ─────────────────────────────────────────────

# def load_freedom_house(filepath):
#     print(f"\nLoading Freedom House: {filepath}")

#     xl = pd.ExcelFile(filepath)

#     target_sheet = xl.sheet_names[1]

#     df = pd.read_excel(filepath, sheet_name=target_sheet, header=1)

#     df.columns = [str(c).strip() for c in df.columns]

#     # filter by edition/year if available
#     year_col = next(
#         (c for c in df.columns if c.lower() == "edition"),
#         None
#     )

#     if year_col:
#         df = df[df[year_col] == FH_YEAR].copy()

#     # detect columns
#     country_col = next(
#         (c for c in df.columns
#          if c.lower() in ["country/territory"]),
#         None
#     )

#     pr_col = next(
#         (c for c in df.columns
#          if c.upper() in ["PR"]),
#         None
#     )

#     cl_col = next(
#         (c for c in df.columns
#          if c.upper() in ["CL", "B", "CIVIL LIBERTIES"]),
#         None
#     )

#     total_col = next(
#         (c for c in df.columns
#          if c.upper() in ["TOTAL"]),
#         None
#     )

#     status_col = next(
#         (c for c in df.columns
#          if c.lower() in ["status"]),
#         None
#     )

#     rename = {}

#     if country_col:
#         rename[country_col] = "country_name"

#     if pr_col:
#         rename[pr_col] = "fh_political_rights"

#     if cl_col:
#         rename[cl_col] = "fh_civil_liberties"

#     if total_col:
#         rename[total_col] = "fh_total_score"

#     if status_col:
#         rename[status_col] = "fh_status"

#     df = df.rename(columns=rename)

#     keep_cols = list(rename.values())

#     df = df[keep_cols].dropna(subset=["country_name"])

#     # normalize names
#     df["country_name"] = (
#         df["country_name"]
#         .str.strip()
#         .replace(COUNTRY_ALIASES)
#     )

#     print(f"Loaded {len(df)} Freedom House rows")

#     return df

def load_freedom_house(filepath):
    print(f"\nLoading Freedom House: {filepath}")

    xl = pd.ExcelFile(filepath)
    target_sheet = xl.sheet_names[1]

    df = pd.read_excel(filepath, sheet_name=target_sheet, header=1)

    df.columns = [str(c).strip() for c in df.columns]

    # filter by edition/year
    if "Edition" in df.columns:
        df = df[df["Edition"] == FH_YEAR].copy()

    # keep only the exact columns we want
    df = df.rename(columns={
        "Country/Territory": "country_name",
        "PR": "fh_political_rights",
        "CL": "fh_civil_liberties",
        "Total": "fh_total_score",
        "Status": "fh_status",
    })

    df = df[
        [
            "country_name",
            "fh_political_rights",
            "fh_civil_liberties",
            "fh_total_score",
            "fh_status",
        ]
    ].copy()

    # normalize names
    df["country_name"] = (
        df["country_name"]
        .str.strip()
        .replace(COUNTRY_ALIASES)
    )

    print(f"Loaded {len(df)} Freedom House rows")

    return df

# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────

def main():

    vdem_df = load_vdem(VDEM_FILE)

    fh_df = load_freedom_house(FH_FILE)

    print("\nMerging datasets...")

    merged = vdem_df.merge(
        fh_df,
        on="country_name",
        how="left"
    )

    print(f"Merged rows: {len(merged)}")

    merged.to_csv(OUTPUT_FILE, index=False)

    print(f"\nSaved merged dataset to: {OUTPUT_FILE}")

    print("\nPreview:")
    print(
        merged[
            [
                "country_name",
                "year",
                "electoral_democracy",
                "liberal_democracy",
                "fh_total_score",
                "fh_status",
            ]
        ].head().to_string(index=False)
    )


if __name__ == "__main__":
    main()