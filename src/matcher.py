"""
Material matching logic for mapping HiCAD materials to Oekobilanz entries.
"""
import json
from pathlib import Path
from dataclasses import dataclass
from typing import Optional


@dataclass
class MaterialMatch:
    """Result of a material matching attempt."""
    matched: bool
    oeko_id: Optional[str] = None
    oeko_name: Optional[str] = None
    ubp_per_kg: Optional[float] = None
    category: Optional[str] = None
    match_type: str = "none"  # "exact", "type_override", "none"


@dataclass
class CoatingMatch:
    """Result of a coating matching attempt."""
    matched: bool
    oeko_id: Optional[str] = None
    oeko_name: Optional[str] = None
    ubp_per_m2: Optional[float] = None


class MaterialMatcher:
    """
    Matches HiCAD material codes to Oekobilanz material entries.
    """

    def __init__(self, mapping_file: str | Path = "data/material_map.json"):
        self.mapping_file = Path(mapping_file)
        self._load_mappings()

    def _load_mappings(self):
        """Load material mappings from JSON file."""
        if self.mapping_file.exists():
            with open(self.mapping_file, "r", encoding="utf-8") as f:
                data = json.load(f)
                self.materials = data.get("materials", {})
                self.type_overrides = data.get("type_overrides", {})
                self.type_fallback = data.get("type_fallback", {})
                self.coatings = data.get("coatings", {})
                self.coating_alu_override = data.get("coating_aluminum_override", {})
        else:
            self.materials = {}
            self.type_overrides = {}
            self.type_fallback = {}
            self.coatings = {}
            self.coating_alu_override = {}

    def _normalize(self, text: str) -> str:
        """Normalize text for matching."""
        return text.strip().upper() if text else ""

    def _is_aluminum(self, material: str) -> bool:
        """Check if material is aluminum-based."""
        mat_upper = material.upper()
        return any(keyword in mat_upper for keyword in
                   ["AL", "ALUMINIUM", "ALUMINUM", "LEICHTMETALL", "AW-6060", "AL99"])

    def _is_steel(self, material: str) -> bool:
        """Check if material is steel-based."""
        mat_upper = material.upper()
        return any(keyword in mat_upper for keyword in
                   ["S235", "S355", "STAHL", "STEEL", "METALL ALLGEMEIN"])

    def _get_base_category(self, material: str) -> str:
        """Get base category (steel, aluminum, stainless, unknown)."""
        mat_upper = material.upper()
        if "X5CRNI" in mat_upper or mat_upper == "304":
            return "stainless_steel"
        if self._is_aluminum(material):
            return "aluminum"
        if self._is_steel(material):
            return "steel"
        return "unknown"

    def _check_type_override(self, material: str, typ: str) -> Optional[MaterialMatch]:
        """
        Check if a type override applies for the given material and type.

        Some materials (e.g., steel) have different UBP values depending on
        whether they're used as profiles or sheets. This method checks if
        the component type triggers an override.

        Note: Only applies to regular steel, NOT stainless steel (304, X5CrNi).

        Args:
            material: The base material code
            typ: The component type (e.g., "Bleche", "Kantblech")

        Returns:
            MaterialMatch if an override applies, None otherwise
        """
        category = self._get_base_category(material)

        for type_key, override_data in self.type_overrides.items():
            if type_key.lower() in typ.lower():
                # Only apply sheet override to regular steel, NOT stainless steel
                if category == "steel" and "steel" in override_data:
                    override = override_data["steel"]
                    return MaterialMatch(
                        matched=True,
                        oeko_id=override["oeko_id"],
                        oeko_name=override["oeko_name"],
                        ubp_per_kg=override["ubp_per_kg"],
                        category="steel_sheet",
                        match_type="type_override"
                    )
        return None

    def _check_type_fallback(self, typ: str, bezeichnung: str) -> Optional[MaterialMatch]:
        """
        Check if a type-based fallback applies when material is unknown.

        Used for fasteners, plastics, and other items where type indicates the category.

        Args:
            typ: The component type (e.g., "Sechskantschrauben", "Scheiben")
            bezeichnung: The description (e.g., "U-Kunststoffplatten")

        Returns:
            MaterialMatch if a fallback applies, None otherwise
        """
        # Check fasteners by type
        fasteners = self.type_fallback.get("fasteners", {})
        if typ and fasteners.get("types"):
            for fastener_type in fasteners["types"]:
                if fastener_type.lower() in typ.lower():
                    return MaterialMatch(
                        matched=True,
                        oeko_id=fasteners["oeko_id"],
                        oeko_name=fasteners["oeko_name"],
                        ubp_per_kg=fasteners["ubp_per_kg"],
                        category="fastener",
                        match_type="type_fallback"
                    )

        # Check fasteners by keywords in description (for items without type)
        if bezeichnung and fasteners.get("keywords"):
            for keyword in fasteners["keywords"]:
                if keyword.lower() in bezeichnung.lower():
                    return MaterialMatch(
                        matched=True,
                        oeko_id=fasteners["oeko_id"],
                        oeko_name=fasteners["oeko_name"],
                        ubp_per_kg=fasteners["ubp_per_kg"],
                        category="fastener",
                        match_type="type_fallback"
                    )

        # Check plastic by keywords in description
        plastic = self.type_fallback.get("plastic", {})
        if bezeichnung and plastic.get("keywords"):
            for keyword in plastic["keywords"]:
                if keyword.lower() in bezeichnung.lower():
                    return MaterialMatch(
                        matched=True,
                        oeko_id=plastic["oeko_id"],
                        oeko_name=plastic["oeko_name"],
                        ubp_per_kg=plastic["ubp_per_kg"],
                        category="plastic",
                        match_type="type_fallback"
                    )

        return None

    def match_material(self, material: str, typ: str = "",
                       bezeichnung: str = "") -> MaterialMatch:
        """
        Match a material from Mengenliste to Oekobilanz entry.

        Args:
            material: Material code (e.g., "S235JR")
            typ: Profile type (e.g., "U - Profile")
            bezeichnung: Description (e.g., "UPE 300")

        Returns:
            MaterialMatch with matched data or unmatched status
        """
        # Clean inputs
        material = material.strip() if material else ""
        typ = typ.strip() if typ else ""
        bezeichnung = bezeichnung.strip() if bezeichnung else ""

        if not material or material == " ":
            # Try type-based fallback for items without material (fasteners, plastics)
            fallback_match = self._check_type_fallback(typ, bezeichnung)
            if fallback_match:
                return fallback_match
            return MaterialMatch(matched=False, match_type="empty")

        # Try exact match first
        if material in self.materials:
            mat_data = self.materials[material]

            # Check for type override (e.g., Bleche should use sheet, not profile)
            override_match = self._check_type_override(material, typ)
            if override_match:
                return override_match

            # Return base match
            return MaterialMatch(
                matched=True,
                oeko_id=mat_data["oeko_id"],
                oeko_name=mat_data["oeko_name"],
                ubp_per_kg=mat_data["ubp_per_kg"],
                category=mat_data.get("category", "unknown"),
                match_type="exact"
            )

        return MaterialMatch(matched=False, match_type="no_mapping")

    def match_coating(self, beschichtung: str, material: str = "") -> CoatingMatch:
        """
        Match a coating to Oekobilanz entry.

        Args:
            beschichtung: Coating description (e.g., "feuerverzinkt")
            material: Base material (for aluminum/steel differentiation)

        Returns:
            CoatingMatch with matched data or unmatched status
        """
        if not beschichtung or beschichtung.strip() == "":
            return CoatingMatch(matched=False)

        beschichtung_lower = beschichtung.lower()
        is_aluminum = self._is_aluminum(material)

        # Try each coating pattern
        for coating_key, coating_data in self.coatings.items():
            if coating_key.lower() in beschichtung_lower:
                # Check for aluminum override
                if is_aluminum and coating_key in self.coating_alu_override:
                    alu_data = self.coating_alu_override[coating_key]
                    return CoatingMatch(
                        matched=True,
                        oeko_id=alu_data["oeko_id"],
                        oeko_name=alu_data["oeko_name"],
                        ubp_per_m2=alu_data["ubp_per_m2"]
                    )
                else:
                    return CoatingMatch(
                        matched=True,
                        oeko_id=coating_data["oeko_id"],
                        oeko_name=coating_data["oeko_name"],
                        ubp_per_m2=coating_data["ubp_per_m2"]
                    )

        return CoatingMatch(matched=False)

    def get_all_mappings(self) -> dict:
        """Return all current mappings for display."""
        return {
            "materials": self.materials,
            "type_overrides": self.type_overrides,
            "coatings": self.coatings,
        }


if __name__ == "__main__":
    # Test matching
    matcher = MaterialMatcher()

    # Test materials
    test_materials = [
        ("S235JR", "U - Profile", "UPE 300"),
        ("S235JR", "Flachstahl", "BRFL 10x360"),
        ("S235JR", "Bleche", "Blech 5mm"),
        ("S235JR", "Kantblech", "Kantblech 3mm"),
        ("EN AW-6060", "", "Klemmprofil"),
        ("X5CrNi18-10", "Bleche", ""),
        ("304", "Kantblech", ""),  # Should NOT be overridden to sheet
        ("EPDM", "", "Dichtung"),
        ("NR", "Schallhemmende Produkte", ""),
        ("PE", "", ""),
        ("", "Sechskantschrauben", "M8x20"),  # Fastener fallback
        ("", "Scheiben", "M8"),  # Fastener fallback
        ("", "", "U-Kunststoffplatten 45x60"),  # Plastic fallback
    ]

    print("=== MATERIAL MATCHING ===")
    for mat, typ, bez in test_materials:
        result = matcher.match_material(mat, typ, bez)
        print(f"{mat:20} | {typ:25} -> {result.oeko_name or 'NO MATCH':30} "
              f"({result.ubp_per_kg or '-'} UBP/kg) [{result.match_type}]")

    # Test coatings
    print("\n=== COATING MATCHING ===")
    test_coatings = [
        ("feuerverzinkt", "S235JR"),
        ("Pulverb. IGP-DURA face", "S235JR"),
        ("Pulverb. IGP-DURA face", "EN AW-6060"),
        ("", "S235JR"),
    ]

    for coating, mat in test_coatings:
        result = matcher.match_coating(coating, mat)
        print(f"{coating[:30]:30} | {mat:15} -> {result.oeko_name or 'NO MATCH':30} "
              f"({result.ubp_per_m2 or '-'} UBP/m2)")
