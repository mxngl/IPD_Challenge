from __future__ import annotations

import csv
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path

from .models import ConstructionItem


CF_TO_CY = 1.0 / 27.0
SOFTWOOD_LUMBER_KG_PER_CF = 12.72
GLULAM_KG_PER_CF = 19.43
CONCRETE_FLOOR_TYPE = "Concrete (sf)"
WOOD_FLOOR_TYPE = "Wood System (sf)"


@dataclass(slots=True)
class StructuralScheduleReport:
    construction_items: list[ConstructionItem]
    mapped_rows: int
    skipped_rows: list[dict[str, str]]

    def to_dict(self) -> dict[str, object]:
        return {
            "mapped_rows": self.mapped_rows,
            "skipped_rows": self.skipped_rows,
            "construction_items": [
                {
                    "assembly": item.assembly,
                    "material_type": item.material_type,
                    "amount": item.amount,
                }
                for item in self.construction_items
            ],
        }


def load_structural_schedule(
    csv_path: Path | str,
) -> StructuralScheduleReport:
    path = Path(csv_path)
    totals: dict[tuple[str, str], float] = defaultdict(float)
    skipped_rows: list[dict[str, str]] = []
    mapped_rows = 0

    with path.open(newline="", encoding="utf-8-sig") as handle:
        reader = csv.DictReader(handle)
        for row in reader:
            mapped = _map_structural_row(row)
            if mapped is None:
                skipped_rows.append(
                    {
                        "element_id": row.get("ElementId", ""),
                        "category": row.get("Category", ""),
                        "family": row.get("Family", ""),
                        "type": row.get("Type", ""),
                        "material": row.get("Material", ""),
                        "volume": row.get("Volume", ""),
                        "reason": "No STV mapping rule matched this row.",
                    }
                )
                continue
            mapped_rows += 1
            totals[(mapped.assembly, mapped.material_type)] += mapped.amount

    construction_items = [
        ConstructionItem(assembly=assembly, material_type=material_type, amount=amount)
        for (assembly, material_type), amount in sorted(totals.items())
        if amount > 0
    ]
    return StructuralScheduleReport(
        construction_items=construction_items,
        mapped_rows=mapped_rows,
        skipped_rows=skipped_rows,
    )


def _map_structural_row(row: dict[str, str]) -> ConstructionItem | None:
    category = _normalized(row.get("Category"))
    family = _normalized(row.get("Family"))
    type_name = _normalized(row.get("Type"))
    material = _normalized(row.get("Material"))
    volume_cf = _parse_measurement(row.get("Volume", ""))
    area_sf = _parse_measurement(row.get("Area", ""))

    if category == "structural columns":
        if "concrete" in family or "concrete" in material:
            return ConstructionItem(
                assembly="Columns",
                material_type="Reinforced Concrete Column (cy)",
                amount=volume_cf * CF_TO_CY,
            )
        if "glulam" in family or "timber" in family:
            return ConstructionItem(
                assembly="Columns",
                material_type="Glulam Column (kg)",
                amount=volume_cf * GLULAM_KG_PER_CF,
            )
        # if "timber" in family or "softwood" in material or "lumber" in material:
        #     return ConstructionItem(
        #         assembly="Columns",
        #         material_type="Wood Column (kg)",
        #         amount=volume_cf * SOFTWOOD_LUMBER_KG_PER_CF,
        #     )
        return None

    if category == "structural framing":
        if "concrete" in family or "concrete" in material:
            return ConstructionItem(
                assembly="Beams",
                material_type="Reinforced Concrete Beam (cy)",
                amount=volume_cf * CF_TO_CY,
            )
        if "glulam" in family:
            return ConstructionItem(
                assembly="Beams",
                material_type="Glulam Beam (kg)",
                amount=volume_cf * GLULAM_KG_PER_CF,
            )
        if "wood" in family or "softwood" in material or "lumber" in material:
            return ConstructionItem(
                assembly="Beams",
                material_type="Wood Beam (kg)",
                amount=volume_cf * SOFTWOOD_LUMBER_KG_PER_CF,
            )
        return None

    if category == "structural foundations":
        if "slab" in family or "slab" in type_name:
            return ConstructionItem(
                assembly="Foundation",
                material_type="Concrete Slab (cy)",
                amount=volume_cf * CF_TO_CY,
            )
        if "footing" in family or "footing" in type_name:
            return ConstructionItem(
                assembly="Foundation",
                material_type="Strip Foundation (cy)",
                amount=volume_cf * CF_TO_CY,
            )
        if "concrete" in material:
            return ConstructionItem(
                assembly="Foundation",
                material_type="Concrete Slab (cy)",
                amount=volume_cf * CF_TO_CY,
            )
        return None

    if category == "floors":
        if area_sf <= 0:
            return None
        if _is_wood_floor(family, type_name, material):
            return ConstructionItem(
                assembly="Floor",
                material_type=WOOD_FLOOR_TYPE,
                amount=area_sf,
            )
        if _is_concrete_floor(family, type_name, material):
            return ConstructionItem(
                assembly="Floor",
                material_type=CONCRETE_FLOOR_TYPE,
                amount=area_sf,
            )
        return None

    return None


def _parse_measurement(raw_value: str) -> float:
    text = (raw_value or "").strip()
    if not text:
        return 0.0
    token = text.split()[0].replace(",", "")
    return float(token)


def _is_wood_floor(family: str, type_name: str, material: str) -> bool:
    search_text = " ".join((family, type_name, material))
    return any(keyword in search_text for keyword in ("wood", "timber", "glulam", "lumber", "clt", "plywood"))


def _is_concrete_floor(family: str, type_name: str, material: str) -> bool:
    search_text = " ".join((family, type_name, material))
    return any(keyword in search_text for keyword in ("concrete", "slab", "deck"))


def _normalized(value: str | None) -> str:
    return (value or "").strip().lower()
