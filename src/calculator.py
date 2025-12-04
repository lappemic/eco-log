"""
UBP calculation logic for environmental impact assessment.
"""
import logging
import math
import pandas as pd
from dataclasses import dataclass, field
from typing import Optional
from .matcher import MaterialMatcher, MaterialMatch, CoatingMatch

logger = logging.getLogger(__name__)


@dataclass
class ComponentResult:
    """UBP calculation result for a single component."""
    pos: float
    anzahl: int
    bezeichnung: str
    material: str
    typ: str
    beschichtung: str
    ges_gewicht_kg: float
    flaeche_m2: float

    # Matching results
    material_matched: bool = False
    material_oeko_name: Optional[str] = None
    material_ubp_per_kg: Optional[float] = None

    coating_matched: bool = False
    coating_oeko_name: Optional[str] = None
    coating_ubp_per_m2: Optional[float] = None

    # Calculated UBP
    ubp_material: float = 0.0
    ubp_coating: float = 0.0
    ubp_total: float = 0.0

    # Reason if not matched
    unmatched_reason: str = ""


@dataclass
class CalculationResults:
    """Complete UBP calculation results."""
    components: list[ComponentResult] = field(default_factory=list)

    # Summary
    total_ubp: float = 0.0
    total_weight_kg: float = 0.0
    total_area_m2: float = 0.0
    components_matched: int = 0
    components_total: int = 0
    match_rate: float = 0.0

    # Aggregations
    by_material: dict = field(default_factory=dict)
    by_coating: dict = field(default_factory=dict)
    by_type: dict = field(default_factory=dict)
    unmatched: list = field(default_factory=list)


class UBPCalculator:
    """
    Calculate UBP (Umweltbelastungspunkte) for model components.
    """

    def __init__(self, matcher: Optional[MaterialMatcher] = None):
        self.matcher = matcher or MaterialMatcher()

    def calculate(self, df: pd.DataFrame) -> CalculationResults:
        """
        Calculate UBP for all components in a DataFrame.

        Args:
            df: DataFrame from parse_mengenliste

        Returns:
            CalculationResults with all calculations and aggregations
        """
        results = CalculationResults()
        results.components_total = len(df)

        for _, row in df.iterrows():
            comp_result = self._calculate_component(row)
            results.components.append(comp_result)

            # Update totals
            results.total_ubp += comp_result.ubp_total
            results.total_weight_kg += comp_result.ges_gewicht_kg
            results.total_area_m2 += comp_result.flaeche_m2

            if comp_result.material_matched:
                results.components_matched += 1

                # Aggregate by material
                mat_name = comp_result.material_oeko_name or "Unknown"
                if mat_name not in results.by_material:
                    results.by_material[mat_name] = {
                        "ubp": 0.0, "weight_kg": 0.0, "count": 0
                    }
                results.by_material[mat_name]["ubp"] += comp_result.ubp_material
                results.by_material[mat_name]["weight_kg"] += comp_result.ges_gewicht_kg
                results.by_material[mat_name]["count"] += 1

                # Aggregate by type
                typ = comp_result.typ or "Unknown"
                if typ not in results.by_type:
                    results.by_type[typ] = {"ubp": 0.0, "count": 0}
                results.by_type[typ]["ubp"] += comp_result.ubp_total
                results.by_type[typ]["count"] += 1
            else:
                results.unmatched.append({
                    "pos": comp_result.pos,
                    "material": comp_result.material,
                    "typ": comp_result.typ,
                    "weight_kg": comp_result.ges_gewicht_kg,
                    "reason": comp_result.unmatched_reason
                })

            # Aggregate coatings
            if comp_result.coating_matched:
                coat_name = comp_result.coating_oeko_name or "Unknown"
                if coat_name not in results.by_coating:
                    results.by_coating[coat_name] = {
                        "ubp": 0.0, "area_m2": 0.0, "count": 0
                    }
                results.by_coating[coat_name]["ubp"] += comp_result.ubp_coating
                results.by_coating[coat_name]["area_m2"] += comp_result.flaeche_m2
                results.by_coating[coat_name]["count"] += 1

        # Calculate match rate
        if results.components_total > 0:
            results.match_rate = results.components_matched / results.components_total

        return results

    def _calculate_component(self, row: pd.Series) -> ComponentResult:
        """Calculate UBP for a single component."""
        # Extract values from row
        pos = row.get("pos", 0)
        anzahl = int(row.get("anzahl", 1))
        bezeichnung = str(row.get("bezeichnung", ""))
        material = str(row.get("material", ""))
        typ = str(row.get("typ", ""))
        beschichtung = str(row.get("beschichtung", ""))

        # Handle NaN values with logging
        raw_weight = row.get("ges_gewicht_kg", 0)
        raw_area = row.get("flaeche_m2", 0)

        if raw_weight is None or (isinstance(raw_weight, float) and math.isnan(raw_weight)):
            logger.warning(f"Pos {pos}: ges_gewicht_kg is NaN, using 0")
            ges_gewicht_kg = 0.0
        else:
            ges_gewicht_kg = float(raw_weight)

        if raw_area is None or (isinstance(raw_area, float) and math.isnan(raw_area)):
            logger.warning(f"Pos {pos}: flaeche_m2 is NaN, using 0")
            flaeche_m2 = 0.0
        else:
            flaeche_m2 = float(raw_area)

        result = ComponentResult(
            pos=pos,
            anzahl=anzahl,
            bezeichnung=bezeichnung,
            material=material,
            typ=typ,
            beschichtung=beschichtung,
            ges_gewicht_kg=ges_gewicht_kg,
            flaeche_m2=flaeche_m2,
        )

        # Match material
        mat_match = self.matcher.match_material(material, typ, bezeichnung)
        if mat_match.matched:
            result.material_matched = True
            result.material_oeko_name = mat_match.oeko_name
            result.material_ubp_per_kg = mat_match.ubp_per_kg
            result.ubp_material = ges_gewicht_kg * (mat_match.ubp_per_kg or 0)
        else:
            result.unmatched_reason = mat_match.match_type

        # Match coating
        coat_match = self.matcher.match_coating(beschichtung, material)
        if coat_match.matched:
            result.coating_matched = True
            result.coating_oeko_name = coat_match.oeko_name
            result.coating_ubp_per_m2 = coat_match.ubp_per_m2
            result.ubp_coating = flaeche_m2 * (coat_match.ubp_per_m2 or 0)

        # Total UBP
        result.ubp_total = result.ubp_material + result.ubp_coating

        return result

    def results_to_dataframe(self, results: CalculationResults) -> pd.DataFrame:
        """Konvertiere Ergebnisse zu DataFrame für Anzeige/Export."""
        data = []
        for comp in results.components:
            data.append({
                "Pos": comp.pos,
                "Anzahl": comp.anzahl,
                "Bezeichnung": comp.bezeichnung,
                "Material": comp.material,
                "Typ": comp.typ,
                "Beschichtung": comp.beschichtung,
                "Gewicht (kg)": round(comp.ges_gewicht_kg, 2),
                "Fläche (m²)": round(comp.flaeche_m2, 2),
                "Material (Oeko)": comp.material_oeko_name or "-",
                "Beschichtung (Oeko)": comp.coating_oeko_name or "-",
                "UBP Material": round(comp.ubp_material, 0),
                "UBP Beschichtung": round(comp.ubp_coating, 0),
                "UBP Total": round(comp.ubp_total, 0),
                "Zugeordnet": "Ja" if comp.material_matched else "Nein",
            })
        return pd.DataFrame(data)

    def summary_to_dict(self, results: CalculationResults) -> dict:
        """Convert summary to dict for JSON/display."""
        return {
            "total_ubp": round(results.total_ubp, 0),
            "total_weight_kg": round(results.total_weight_kg, 2),
            "total_area_m2": round(results.total_area_m2, 2),
            "components_matched": results.components_matched,
            "components_total": results.components_total,
            "match_rate": round(results.match_rate * 100, 1),
            "by_material": {
                k: {"ubp": round(v["ubp"], 0), "weight_kg": round(v["weight_kg"], 2)}
                for k, v in results.by_material.items()
            },
            "by_coating": {
                k: {"ubp": round(v["ubp"], 0), "area_m2": round(v["area_m2"], 2)}
                for k, v in results.by_coating.items()
            },
            "unmatched_count": len(results.unmatched),
        }


if __name__ == "__main__":
    # Test calculation with real data
    import glob
    from .parser import parse_mengenliste

    model_files = glob.glob("25-008-C*.xlsx")
    if model_files:
        print("=== LOADING MODEL ===")
        df = parse_mengenliste(model_files[0])
        print(f"Loaded {len(df)} components")

        print("\n=== CALCULATING UBP ===")
        calculator = UBPCalculator()
        results = calculator.calculate(df)

        summary = calculator.summary_to_dict(results)
        print(f"\nTotal UBP: {summary['total_ubp']:,.0f}")
        print(f"Total Weight: {summary['total_weight_kg']:,.2f} kg")
        print(f"Match Rate: {summary['match_rate']}%")

        print("\n=== BY MATERIAL ===")
        for mat, data in sorted(results.by_material.items(),
                                key=lambda x: -x[1]["ubp"]):
            pct = data["ubp"] / results.total_ubp * 100 if results.total_ubp > 0 else 0
            print(f"{mat:35} {data['ubp']:>12,.0f} UBP ({pct:>5.1f}%)")

        print("\n=== BY COATING ===")
        for coat, data in sorted(results.by_coating.items(),
                                 key=lambda x: -x[1]["ubp"]):
            print(f"{coat:35} {data['ubp']:>12,.0f} UBP")

        print(f"\n=== UNMATCHED ({len(results.unmatched)}) ===")
        for item in results.unmatched[:10]:
            print(f"Pos {item['pos']}: {item['material']} - {item['typ']} ({item['reason']})")
