from __future__ import annotations

import csv
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path

from .models import ConstructionItem


DOOR_DEFAULT_THICKNESS_FT = 1.75 / 12.0


@dataclass(slots=True)
class ArchitectureScheduleReport:
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


def load_architecture_schedule(csv_path: Path | str) -> ArchitectureScheduleReport:
    path = Path(csv_path)
    totals: dict[tuple[str, str], float] = defaultdict(float)
    skipped_rows: list[dict[str, str]] = []
    mapped_rows = 0

    with path.open(newline="", encoding="utf-8-sig") as handle:
        reader = csv.DictReader(handle)
        for row in reader:
            mapped = _map_architecture_row(row)
            if mapped is None:
                skipped_rows.append(
                    {
                        "element_id": row.get("ElementId", ""),
                        "category": row.get("Category", ""),
                        "family": row.get("Family", ""),
                        "type": row.get("Type", ""),
                        "assembly_code": row.get("Assembly Code", ""),
                        "assembly_description": row.get("Assembly Description", ""),
                        "material": row.get("Material", ""),
                        "area": row.get("Area", ""),
                        "volume": row.get("Volume", ""),
                        "reason": "No STV architecture mapping rule matched this row.",
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

    return ArchitectureScheduleReport(
        construction_items=construction_items,
        mapped_rows=mapped_rows,
        skipped_rows=skipped_rows,
    )


def _map_architecture_row(row: dict[str, str]) -> ConstructionItem | None:
    category = _normalized(row.get("Category"))
    family = _normalized(row.get("Family"))
    type_name = _normalized(row.get("Type"))
    assembly_code = _normalized(row.get("Assembly Code"))
    assembly_description = _normalized(row.get("Assembly Description"))
    material = _normalized(row.get("Material"))
    area_sf = _parse_measurement(row.get("Area", ""))
    volume_cf = _parse_measurement(row.get("Volume", ""))

    if category == "floors":
        return _map_floor(area_sf, family, type_name, assembly_code, assembly_description, material)

    if category == "walls":
        return _map_wall(area_sf, family, type_name, assembly_code, assembly_description, material)

    if category == "curtain panels":
        if area_sf <= 0:
            return None
        return ConstructionItem(
            assembly="Exterior Wall",
            material_type="Curtain Wall Double Pane (sf)",
            amount=area_sf,
        )
    
    if category == "curtain wall mullions":
        if area_sf <= 0:
            return None
        return ConstructionItem(
            assembly="Exterior Wall",
            material_type="Curtain Wall Double Pane (sf)",
            amount=area_sf,
        )

    if category == "roofs":
        if area_sf <= 0:
            return None
        if _contains_any(type_name, material, "green"):
            material_type = "Green Roof (sf)"
        elif _contains_any(type_name, material, "epdm", "membrane"):
            material_type = "EPDM Membrane (sf)"
        elif _contains_any(type_name, material, "wood", "timber", "lumber"):
            material_type = "Wood Structure (sf)"
        else:
            material_type = "EPDM Membrane (sf)"
        return ConstructionItem(assembly="Roof", material_type=material_type, amount=area_sf)

    if category == "doors":
        amount_sf = area_sf if area_sf > 0 else _door_area_fallback(row)
        if amount_sf <= 0:
            return None
        return ConstructionItem(
            assembly="Interior Wall",
            material_type="Timber Studs and Painted Gypsum (sf)",
            amount=amount_sf,
        )

    return None


def _map_floor(
    area_sf: float,
    family: str,
    type_name: str,
    assembly_code: str,
    assembly_description: str,
    material: str,
) -> ConstructionItem | None:
    if area_sf <= 0:
        return None

    search_text = " ".join((family, type_name, assembly_code, assembly_description, material))
    if _contains_any(search_text, "wood", "timber", "glulam", "clt", "lumber"):
        material_type = "Wood System (sf)"
    else:
        material_type = "Concrete (sf)"

    return ConstructionItem(assembly="Floor", material_type=material_type, amount=area_sf)


def _map_wall(
    area_sf: float,
    family: str,
    type_name: str,
    assembly_code: str,
    assembly_description: str,
    material: str,
) -> ConstructionItem | None:
    if area_sf <= 0:
        return None

    search_text = " ".join((family, type_name, assembly_code, assembly_description, material))
    is_exterior = (
        assembly_code.startswith("b20")
        or _contains_any(search_text, "exterior", "storefront", "curtain")
    )

    if is_exterior:
        if _contains_any(search_text, "storefront", "curtain", "glazing", "glass"):
            material_type = "Curtain Wall Double Pane (sf)"
        elif _contains_any(search_text, "brick"):
            material_type = "Brick on Metal Stud (sf)"
        elif _contains_any(search_text, "concrete"):
            material_type = "Concrete Cladding (sf)"
        else:
            material_type = "EIFS on Metal Stud (sf)"
        return ConstructionItem(
            assembly="Exterior Wall",
            material_type=material_type,
            amount=area_sf,
        )

    if _contains_any(search_text, "concrete"):
        material_type = "Concrete and Painted Gypsum (sf)"
    elif _contains_any(search_text, "wood", "timber", "lumber"):
        material_type = "Timber Studs and Painted Gypsum (sf)"
    else:
        material_type = "Steel Studs and Painted Gypsum (sf)"

    return ConstructionItem(
        assembly="Interior Wall",
        material_type=material_type,
        amount=area_sf,
    )


def _door_area_fallback(row: dict[str, str]) -> float:
    width_ft = _parse_length_feet(row.get("Width", ""))
    height_ft = _parse_length_feet(row.get("Height", ""))
    if width_ft > 0 and height_ft > 0:
        return width_ft * height_ft

    volume_cf = _parse_measurement(row.get("Volume", ""))
    if volume_cf > 0 and width_ft > 0:
        return volume_cf / DOOR_DEFAULT_THICKNESS_FT

    return 0.0


def _contains_any(*values: str) -> bool:
    if len(values) < 2:
        return False
    text = values[0]
    keywords = values[1:]
    return any(keyword in text for keyword in keywords)


def _parse_measurement(raw_value: str) -> float:
    text = (raw_value or "").strip()
    if not text:
        return 0.0
    token = text.split()[0].replace(",", "")
    return float(token)


def _parse_length_feet(raw_value: str | None) -> float:
    text = (raw_value or "").strip()
    if not text:
        return 0.0

    feet = 0.0
    inches = 0.0

    if "'" in text:
        feet_part = text.split("'", 1)[0].strip()
        if feet_part:
            feet = float(feet_part)

    if '"' in text:
        inches_part = text.split('"', 1)[0].split("'")[-1].replace("-", "").strip()
        if inches_part:
            inches = _parse_inches_fraction(inches_part)

    if feet == 0.0 and inches == 0.0:
        return _parse_measurement(text)
    return feet + inches / 12.0


def _parse_inches_fraction(text: str) -> float:
    value = text.strip()
    if not value:
        return 0.0
    if " " in value:
        whole, fraction = value.split(" ", 1)
        return float(whole) + _parse_fraction(fraction)
    if "/" in value:
        return _parse_fraction(value)
    return float(value)


def _parse_fraction(value: str) -> float:
    numerator, denominator = value.split("/", 1)
    denominator_value = float(denominator)
    if denominator_value == 0:
        return 0.0
    return float(numerator) / denominator_value


def _normalized(value: str | None) -> str:
    return (value or "").strip().lower()
