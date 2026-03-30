from __future__ import annotations

import csv
import math
import re
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from .models import ConstructionItem


@dataclass(slots=True)
class MEPScheduleReport:
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


def load_mep_schedule(csv_path: Path | str) -> MEPScheduleReport:
    path = Path(csv_path)
    totals: dict[tuple[str, str], float] = defaultdict(float)
    skipped_rows: list[dict[str, str]] = []
    mapped_rows = 0

    with path.open(newline="", encoding="utf-8-sig") as handle:
        reader = csv.DictReader(handle)
        for row in reader:
            mapped = _map_mep_row(row)
            if mapped is None:
                skipped_rows.append(
                    {
                        "element_id": row.get("ElementId", ""),
                        "category": row.get("Category", ""),
                        "family": row.get("Family", ""),
                        "type": row.get("Type", ""),
                        "system_type": row.get("System Type", ""),
                        "size": row.get("Size", ""),
                        "length": row.get("Length", ""),
                        "area": row.get("Area", ""),
                        "volume": row.get("Volume", ""),
                        "material": row.get("Material", ""),
                        "reason": "No STV MEP mapping rule matched this row.",
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

    return MEPScheduleReport(
        construction_items=construction_items,
        mapped_rows=mapped_rows,
        skipped_rows=skipped_rows,
    )


def _map_mep_row(row: dict[str, str]) -> ConstructionItem | None:
    family_category_key = _family_category_key(row)
    if family_category_key:
        explicit_handler = MEP_FAMILY_CATEGORY_MAPPINGS.get(family_category_key)
        if explicit_handler is not None:
            return explicit_handler(row)

    category = _normalized(row.get("Category"))
    family = _normalized(row.get("Family"))
    material = _normalized(row.get("Material"))
    snapshot = row.get("Parameter Snapshot", "")

    if category in {"ducts", "flex ducts", "duct fittings", "air terminals"}:
        return _map_steel_duct_from_geometry(row)

    if category in {"pipe fittings", "pipe accessories", "pipes", "flex pipes"}:
        material_type = _resolve_pipe_material_type(material, family, snapshot)
        if not material_type:
            return None
        weight_kg = _resolve_weight_kg(row)
        if weight_kg <= 0:
            return None
        return ConstructionItem(
            assembly="MEP",
            material_type=material_type,
            amount=weight_kg,
        )

    return None


def _family_category_key(row: dict[str, str]) -> str:
    family = (row.get("Family") or "").strip()
    category = (row.get("Category") or "").strip()
    if not family or not category:
        return ""
    return f"{family} : {category}"


def _map_steel_duct_from_geometry(row: dict[str, str]) -> ConstructionItem | None:
    material = _normalized(row.get("Material"))
    if "stainless" in material:
        weight_kg = _resolve_weight_kg(row)
        if weight_kg > 0:
            return ConstructionItem(
                assembly="MEP",
                material_type="Stainless Steel Duct (kg)",
                amount=weight_kg,
            )

    equivalent_length_ft = _resolve_duct_equivalent_length(row)
    if equivalent_length_ft <= 0:
        return None

    nominal_diameter_in = _resolve_nominal_diameter_inches(row)
    material_type = "Steel Duct 12\"D (ft)" if nominal_diameter_in <= 15 else "Steel Duct 18\"D (ft)"
    return ConstructionItem(
        assembly="MEP",
        material_type=material_type,
        amount=equivalent_length_ft,
    )


def _map_air_filter_from_flow(row: dict[str, str]) -> ConstructionItem | None:
    snapshot = row.get("Parameter Snapshot", "")
    airflow_m3s = _resolve_airflow_m3s(row, snapshot)
    if airflow_m3s <= 0:
        return None
    return ConstructionItem(
        assembly="MEP",
        material_type="Air Filters (m^3/s)",
        amount=airflow_m3s,
    )


def _map_ahu_from_flow(row: dict[str, str]) -> ConstructionItem | None:
    snapshot = row.get("Parameter Snapshot", "")
    airflow_m3s = _resolve_airflow_m3s(row, snapshot)
    if airflow_m3s <= 0:
        return None
    return ConstructionItem(
        assembly="MEP",
        material_type="Air Handling Unit (m^3/s)",
        amount=airflow_m3s,
    )


def _map_stainless_duct_from_weight(row: dict[str, str]) -> ConstructionItem | None:
    return _map_stainless_duct_from_weight_with_multiplier(row, surface_multiplier=1.0)


def _map_stainless_duct_elbow_from_weight(row: dict[str, str]) -> ConstructionItem | None:
    return _map_stainless_duct_from_weight_with_multiplier(row, surface_multiplier=1.3)


def _map_stainless_duct_tee_from_weight(row: dict[str, str]) -> ConstructionItem | None:
    return _map_stainless_duct_from_weight_with_multiplier(row, surface_multiplier=1.6)


def _map_stainless_duct_cross_from_weight(row: dict[str, str]) -> ConstructionItem | None:
    return _map_stainless_duct_from_weight_with_multiplier(row, surface_multiplier=1.8)


def _map_stainless_duct_transition_from_weight(row: dict[str, str]) -> ConstructionItem | None:
    return _map_stainless_duct_from_weight_with_multiplier(row, surface_multiplier=1.2)


def _map_stainless_duct_rect_to_round_transition_from_weight(
    row: dict[str, str],
) -> ConstructionItem | None:
    return _map_stainless_duct_from_weight_with_multiplier(row, surface_multiplier=1.25)


def _map_stainless_duct_from_weight_with_multiplier(
    row: dict[str, str], *, surface_multiplier: float
) -> ConstructionItem | None:
    weight_kg = _resolve_weight_kg(row)
    if weight_kg <= 0:
        weight_kg = _estimate_rectangular_duct_weight_kg(row, surface_multiplier=surface_multiplier)
    if weight_kg <= 0:
        return None
    return ConstructionItem(
        assembly="MEP",
        material_type="Stainless Steel Duct (kg)",
        amount=weight_kg,
    )


MEP_FAMILY_CATEGORY_MAPPINGS: dict[str, Callable[[dict[str, str]], ConstructionItem | None]] = {
    "Supply Diffuser : Air Terminals": _map_ahu_from_flow,
    "Exhaust Grill : Air Terminals": _map_ahu_from_flow,
    "PRICE-40FF- Filter Frame Stamped Residential Grille-RETURN Hosted : Air Terminals": _map_air_filter_from_flow,
    "34274 : Electrical Fixtures": lambda row: None,
    "Return Diffuser : Air Terminals": _map_ahu_from_flow,
    "Utility Switchboard : Electrical Equipment": lambda row: None,
    "Outdoor AHU - Horizontal : Mechanical Equipment": _map_ahu_from_flow,
    "Rectangular Elbow - Mitered : Duct Fittings": _map_stainless_duct_elbow_from_weight,
    "Rectangular Tee : Duct Fittings": _map_stainless_duct_tee_from_weight,
    "Rectangular Cross : Duct Fittings": _map_stainless_duct_cross_from_weight,
    "Rectangular Transition - Angle : Duct Fittings": _map_stainless_duct_transition_from_weight,
    "Rectangular to Round Transition - Angle : Duct Fittings": _map_stainless_duct_rect_to_round_transition_from_weight,
}


def _resolve_pipe_material_type(material: str, family: str, snapshot: str) -> str | None:
    search_text = " ".join([material, family, snapshot.lower()])
    if "copper" in search_text:
        return "Copper Pipe (kg)"
    if "stainless" in search_text or "steel" in search_text:
        return "Stainless Steel Pipe (kg)"
    if "hdpe" in search_text or "polyethylene" in search_text:
        return "HDPE Pipe (kg)"
    return None


def _resolve_duct_equivalent_length(row: dict[str, str]) -> float:
    length_ft = _parse_length_feet(row.get("Length", ""))
    if length_ft > 0:
        return length_ft

    area_sf = _parse_measurement_value(row.get("Area", ""))
    volume_cf = _parse_measurement_value(row.get("Volume", ""))
    if area_sf > 0 and volume_cf > 0:
        return volume_cf / area_sf
    if volume_cf > 0:
        return volume_cf ** (1.0 / 3.0)
    return 0.0


def _resolve_nominal_diameter_inches(row: dict[str, str]) -> float:
    for field in ("Diameter", "Size", "Width", "Height"):
        candidate = row.get(field, "")
        diameter = _extract_inches(candidate)
        if diameter > 0:
            return diameter

    snapshot = row.get("Parameter Snapshot", "")
    hydraulic = _extract_snapshot_value(snapshot, "Hydraulic Diameter")
    if hydraulic:
        diameter = _extract_inches(hydraulic)
        if diameter > 0:
            return diameter

    width = _extract_snapshot_value(snapshot, "Duct Width")
    height = _extract_snapshot_value(snapshot, "Duct Height")
    width_in = _extract_inches(width)
    height_in = _extract_inches(height)
    if width_in > 0 and height_in > 0:
        return 2 * width_in * height_in / (width_in + height_in)
    if width_in > 0:
        return width_in
    if height_in > 0:
        return height_in
    return 12.0


def _resolve_weight_kg(row: dict[str, str]) -> float:
    for field in ("Weight", "Unit Weight"):
        value = _parse_measurement_value(row.get(field, ""))
        if value > 0:
            return value
    return 0.0


def _estimate_rectangular_duct_weight_kg(
    row: dict[str, str], *, surface_multiplier: float = 1.0
) -> float:
    width_m, height_m = _resolve_rectangular_dimensions_m(row)
    length_m = _resolve_length_m(row)
    if width_m <= 0 or height_m <= 0 or length_m <= 0:
        return 0.0

    thickness_m = _resolve_duct_thickness_m(width_m, height_m)
    density_kg_per_m3 = 8000.0
    sheet_area_m2 = surface_multiplier * 2.0 * (width_m + height_m) * length_m
    return sheet_area_m2 * thickness_m * density_kg_per_m3


def _resolve_rectangular_dimensions_m(row: dict[str, str]) -> tuple[float, float]:
    width_in = _extract_inches(row.get("Width", ""))
    height_in = _extract_inches(row.get("Height", ""))
    snapshot = row.get("Parameter Snapshot", "")

    if width_in <= 0:
        width_in = _extract_inches(_extract_snapshot_value(snapshot, "Duct Width"))
    if height_in <= 0:
        height_in = _extract_inches(_extract_snapshot_value(snapshot, "Duct Height"))

    if width_in <= 0 or height_in <= 0:
        size = row.get("Size", "")
        parsed_width_in, parsed_height_in = _extract_size_pair_inches(size)
        if width_in <= 0:
            width_in = parsed_width_in
        if height_in <= 0:
            height_in = parsed_height_in

    return width_in * 0.0254, height_in * 0.0254


def _resolve_length_m(row: dict[str, str]) -> float:
    length_ft = _parse_length_feet(row.get("Length", ""))
    if length_ft > 0:
        return length_ft * 0.3048

    snapshot = row.get("Parameter Snapshot", "")
    for key in ("Length", "Duct Length", "Computed Length", "Length 1", "Duct Length 1"):
        candidate_ft = _parse_length_feet(_extract_snapshot_value(snapshot, key))
        if candidate_ft > 0:
            return candidate_ft * 0.3048

    volume_cf = _parse_measurement_value(row.get("Volume", ""))
    width_m, height_m = _resolve_rectangular_dimensions_m(row)
    if volume_cf > 0 and width_m > 0 and height_m > 0:
        volume_m3 = volume_cf * 0.0283168
        cross_section_m2 = width_m * height_m
        if cross_section_m2 > 0:
            return volume_m3 / cross_section_m2

    return 0.0


def _resolve_duct_thickness_m(width_m: float, height_m: float) -> float:
    largest_dimension_m = max(width_m, height_m)
    if largest_dimension_m <= 0.3:
        return 0.0005
    if largest_dimension_m <= 0.6:
        return 0.0006
    return 0.0008


def _extract_size_pair_inches(text: str | None) -> tuple[float, float]:
    raw = (text or "").strip()
    if not raw:
        return 0.0, 0.0

    matches = re.findall(r"(\d+(?:\.\d+)?)\s*\"", raw)
    if len(matches) >= 2:
        return float(matches[0]), float(matches[1])

    metric_matches = re.findall(r"(\d+(?:\.\d+)?)", raw)
    if len(metric_matches) >= 2:
        return float(metric_matches[0]), float(metric_matches[1])

    return 0.0, 0.0


def _resolve_airflow_m3s(row: dict[str, str], snapshot: str) -> float:
    for field in ("Airflow", "Flow"):
        value = _parse_flow_m3s(row.get(field, ""))
        if value > 0:
            return value

    for key in (
        "Supply Air Outlet Flow",
        "Supply Air Inlet Flow",
        "Return Air Inlet Flow",
        "Flow",
    ):
        value = _parse_flow_m3s(_extract_snapshot_value(snapshot, key))
        if value > 0:
            return value

    connector_flow = _parse_measurement_value(row.get("Connector Flow", ""))
    if connector_flow > 0:
        # Connector flow is used as a fallback only. In this dataset it is
        # lower fidelity than the explicit snapshot airflow values.
        return connector_flow

    return 0.0


def _parse_flow_m3s(raw_value: str | None) -> float:
    text = (raw_value or "").strip().lower()
    if not text:
        return 0.0

    value = _parse_measurement_value(text)
    if value <= 0:
        return 0.0

    if "/h" in text:
        return value / 3600.0
    if "/s" in text:
        return value
    return value


def _extract_snapshot_value(snapshot: str, parameter_name: str) -> str:
    prefix = f"{parameter_name}="
    for part in snapshot.split(" | "):
        if part.startswith(prefix):
            return part[len(prefix):]
    return ""


def _extract_inches(text: str | None) -> float:
    raw = (text or "").strip()
    if not raw:
        return 0.0

    values = re.findall(r"(\d+(?:\.\d+)?)\s*\"", raw)
    if values:
        return max(float(value) for value in values)

    metric = re.findall(r"(\d+(?:\.\d+)?)", raw)
    if metric:
        return max(float(value) for value in metric)
    return 0.0


def _parse_measurement_value(raw_value: str | None) -> float:
    text = (raw_value or "").strip()
    if not text:
        return 0.0
    match = re.search(r"-?\d+(?:\.\d+)?", text.replace(",", ""))
    return float(match.group(0)) if match else 0.0


def _parse_length_feet(raw_value: str | None) -> float:
    text = (raw_value or "").strip()
    if not text:
        return 0.0

    feet_match = re.search(r"(-?\d+)\s*'", text)
    inches_match = re.search(r"(-?\d+(?:\s+\d+/\d+|\.\d+)?)\s*\"", text)

    feet = float(feet_match.group(1)) if feet_match else 0.0
    inches = _parse_inches_fraction(inches_match.group(1)) if inches_match else 0.0
    return feet + inches / 12.0


def _parse_inches_fraction(text: str) -> float:
    value = text.strip()
    if " " in value:
        whole, fraction = value.split(" ", 1)
        return float(whole) + _parse_fraction(fraction)
    if "/" in value:
        return _parse_fraction(value)
    return float(value)


def _parse_fraction(value: str) -> float:
    numerator, denominator = value.split("/", 1)
    denominator_value = float(denominator)
    if math.isclose(denominator_value, 0.0):
        return 0.0
    return float(numerator) / denominator_value


def _normalized(value: str | None) -> str:
    return (value or "").strip().lower()
