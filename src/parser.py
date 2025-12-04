"""
Excel parser for HiCAD model exports and Oekobilanz data.
"""
import logging
import pandas as pd
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)


def parse_mengenliste(file_path: str | Path) -> pd.DataFrame:
    """
    Parse the Mengenliste sheet from a HiCAD model export.

    Args:
        file_path: Path to the Excel file

    Returns:
        DataFrame with columns: pos, anzahl, bezeichnung, laenge_mm, breite_mm,
        material, typ, benennung, beschichtung, flaeche_m2, gewicht_kg, ges_gewicht_kg
    """
    # Read the Mengenliste sheet, header is at row 8 (0-indexed: row 7)
    df = pd.read_excel(
        file_path,
        sheet_name="Mengenliste",
        header=7,  # Row 8 has headers
        engine="openpyxl"
    )

    # Rename columns to standardized names
    column_map = {
        "Pos.": "pos",
        "Anzahl": "anzahl",
        "Bezeichnung": "bezeichnung",
        "Länge (mm)": "laenge_mm",
        "Breite (mm)": "breite_mm",
        "Material": "material",
        "Typ": "typ",
        "Benennung": "benennung",
        "Beschichtung": "beschichtung",
        "Fl. (m²)": "flaeche_m2",
        "Gew. (kg)": "gewicht_kg",
        "Ges.gew.": "ges_gewicht_kg",
    }

    # Only keep columns that exist
    existing_cols = {k: v for k, v in column_map.items() if k in df.columns}
    df = df.rename(columns=existing_cols)

    # Keep only renamed columns
    df = df[[col for col in existing_cols.values() if col in df.columns]]

    # Drop rows where pos is NaN (metadata rows)
    df = df.dropna(subset=["pos"])

    # Convert numeric columns with validation logging
    numeric_cols = ["pos", "anzahl", "laenge_mm", "breite_mm",
                    "flaeche_m2", "gewicht_kg", "ges_gewicht_kg"]
    for col in numeric_cols:
        if col in df.columns:
            original = df[col]
            df[col] = pd.to_numeric(df[col], errors="coerce")
            failed_count = df[col].isna().sum() - original.isna().sum()
            if failed_count > 0:
                logger.warning(
                    f"Column '{col}': {failed_count} values could not be converted to numeric"
                )

    # Clean string columns
    string_cols = ["bezeichnung", "material", "typ", "benennung", "beschichtung"]
    for col in string_cols:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()

    # Reset index
    df = df.reset_index(drop=True)

    return df


def parse_oekobilanz(file_path: str | Path) -> pd.DataFrame:
    """
    Parse the Baumaterialien Matériaux sheet from Oekobilanz data.

    Args:
        file_path: Path to the Excel file

    Returns:
        DataFrame with columns: id, name, unit, ubp_total, ubp_herstellung, ubp_entsorgung
    """
    # Read the sheet, headers span multiple rows (4-9), data starts at row 10
    df = pd.read_excel(
        file_path,
        sheet_name="Baumaterialien Matériaux",
        header=None,  # We'll handle headers manually
        skiprows=10,  # Skip header rows, data starts at row 11 (0-indexed: 10)
        engine="openpyxl"
    )

    # Select and rename relevant columns based on analysis
    # Column indices (0-based):
    # 0: ID-Nummer
    # 2: BAUMATERIALIEN (name)
    # 6: Bezug (unit)
    # 7: UBP'21 Total
    # 8: UBP'21 Herstellung
    # 9: UBP'21 Entsorgung

    df = df.iloc[:, [0, 2, 6, 7, 8, 9]]
    df.columns = ["id", "name", "unit", "ubp_total", "ubp_herstellung", "ubp_entsorgung"]

    # Drop rows without valid ID
    df = df.dropna(subset=["id"])
    df = df[df["id"].astype(str).str.strip() != ""]

    # Filter out category headers (ID = 0, or ID contains only digits without dots)
    df = df[df["id"].astype(str).str.contains(r"\.", regex=True, na=False)]

    # Convert numeric columns
    for col in ["ubp_total", "ubp_herstellung", "ubp_entsorgung"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    # Clean string columns
    df["id"] = df["id"].astype(str).str.strip()
    df["name"] = df["name"].fillna("").astype(str).str.strip()
    df["unit"] = df["unit"].fillna("").astype(str).str.strip()

    df = df.reset_index(drop=True)

    return df


def get_oekobilanz_lookup(df: pd.DataFrame) -> dict:
    """
    Create a lookup dictionary from Oekobilanz DataFrame.

    Args:
        df: DataFrame from parse_oekobilanz

    Returns:
        Dict mapping ID to material data
    """
    lookup = {}
    for _, row in df.iterrows():
        lookup[row["id"]] = {
            "name": row["name"],
            "unit": row["unit"],
            "ubp_total": row["ubp_total"],
            "ubp_herstellung": row["ubp_herstellung"],
            "ubp_entsorgung": row["ubp_entsorgung"],
        }
    return lookup


if __name__ == "__main__":
    # Test parsing
    import glob

    # Find and parse model file
    model_files = glob.glob("25-008-C*.xlsx")
    if model_files:
        print("=== MENGENLISTE ===")
        df_model = parse_mengenliste(model_files[0])
        print(f"Rows: {len(df_model)}")
        print(df_model.head(10))
        print("\nColumns:", df_model.columns.tolist())
        print("\nUnique materials:", df_model["material"].unique())

    # Parse Oekobilanz
    oeko_file = "Oekobilanzdaten_ Baubereich_Donne_ecobilans_construction_2009-1-2022_v7.0.xlsx"
    print("\n=== OEKOBILANZ ===")
    df_oeko = parse_oekobilanz(oeko_file)
    print(f"Rows: {len(df_oeko)}")
    print(df_oeko.head(10))

    # Test lookup
    lookup = get_oekobilanz_lookup(df_oeko)
    print(f"\nLookup entries: {len(lookup)}")
    print("Sample entry (06.012):", lookup.get("06.012"))
