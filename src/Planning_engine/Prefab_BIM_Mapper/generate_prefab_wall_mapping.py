from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd


BASE_DIR = Path(__file__).resolve().parent
OUTPUTS_DIR = BASE_DIR / "outputs"
PLANNING_ENGINE_DIR = BASE_DIR.parent
PROJECT_DIR = PLANNING_ENGINE_DIR.parent.parent

CENTRAL_BIM_WITH_TAKT_PATH = PROJECT_DIR / "outputs" / "takt_zones" / "central_bim_model_with_takt.csv"
PREFAB_MAPPING_OUTPUT_PATH = OUTPUTS_DIR / "Prefab_Wall_Mapping.csv"

HOST_CATEGORIES = {"Walls"}
ATTACHED_CATEGORIES = {"Curtain Panels", "Curtain Wall Mullions"}
EXTERIOR_WALL_TYPE_TOKEN = "EXTERIOR"
STOREFRONT_WALL_TYPE_TOKEN = "STOREFRONT"
BBOX_TOLERANCE_FT = 0.5


def load_central_bim() -> pd.DataFrame:
    frame = pd.read_csv(CENTRAL_BIM_WITH_TAKT_PATH)
    frame.columns = [str(column).strip() for column in frame.columns]
    return frame


def normalized_level_token(level: object) -> str:
    text = str(level).strip()
    if not text:
        return "UNASSIGNED"
    return text.replace(" ", "").replace("-", "NEG")


def build_prefab_wall_mapping(
    bim_df: pd.DataFrame,
    bbox_tolerance_ft: float = BBOX_TOLERANCE_FT,
) -> pd.DataFrame:
    required_bbox_columns = [
        "Bounding Box Min X (ft)",
        "Bounding Box Min Y (ft)",
        "Bounding Box Min Z (ft)",
        "Bounding Box Max X (ft)",
        "Bounding Box Max Y (ft)",
        "Bounding Box Max Z (ft)",
    ]
    if bim_df.empty or any(column not in bim_df.columns for column in required_bbox_columns):
        return pd.DataFrame(
            columns=[
                "prefab_group_id",
                "host_wall_element_id",
                "element_id",
                "category",
                "family",
                "type",
                "level",
                "takt_id",
                "source_schedule",
            ]
        )

    enriched = bim_df.copy()
    enriched["element_center_x_ft"] = enriched["Bounding Box Center X (ft)"].where(
        enriched["Bounding Box Center X (ft)"].notna(),
        enriched["Position X (ft)"],
    )
    enriched["element_center_y_ft"] = enriched["Bounding Box Center Y (ft)"].where(
        enriched["Bounding Box Center Y (ft)"].notna(),
        enriched["Position Y (ft)"],
    )
    enriched["element_center_z_ft"] = enriched["Bounding Box Center Z (ft)"].where(
        enriched["Bounding Box Center Z (ft)"].notna(),
        enriched["Position Z (ft)"],
    )

    exterior_walls = enriched[
        enriched["Category"].astype(str).isin(HOST_CATEGORIES)
        & enriched["Type"].astype(str).str.contains(EXTERIOR_WALL_TYPE_TOKEN, case=False, na=False)
    ].copy()
    exterior_walls = exterior_walls.dropna(
        subset=[
            "element_center_x_ft",
            "element_center_y_ft",
            "element_center_z_ft",
            *required_bbox_columns,
        ]
    )
    exterior_walls = exterior_walls.sort_values(
        ["Level", "element_center_y_ft", "element_center_x_ft", "ElementId"]
    ).reset_index(drop=True)

    generic_host_walls = exterior_walls[
        ~exterior_walls["Type"].astype(str).str.contains(STOREFRONT_WALL_TYPE_TOKEN, case=False, na=False)
    ].copy()
    storefront_walls = exterior_walls[
        exterior_walls["Type"].astype(str).str.contains(STOREFRONT_WALL_TYPE_TOKEN, case=False, na=False)
    ].copy()

    nested_storefront_ids: set[str] = set()
    for _, storefront in storefront_walls.iterrows():
        level_text = str(storefront.get("Level", "")).strip()
        center_x = float(storefront["element_center_x_ft"])
        center_y = float(storefront["element_center_y_ft"])
        center_z = float(storefront["element_center_z_ft"])
        generic_matches = generic_host_walls[
            (generic_host_walls["Level"].fillna("").astype(str).str.strip() == level_text)
            & (generic_host_walls["Bounding Box Min X (ft)"] - bbox_tolerance_ft <= center_x)
            & (generic_host_walls["Bounding Box Max X (ft)"] + bbox_tolerance_ft >= center_x)
            & (generic_host_walls["Bounding Box Min Y (ft)"] - bbox_tolerance_ft <= center_y)
            & (generic_host_walls["Bounding Box Max Y (ft)"] + bbox_tolerance_ft >= center_y)
            & (generic_host_walls["Bounding Box Min Z (ft)"] - bbox_tolerance_ft <= center_z)
            & (generic_host_walls["Bounding Box Max Z (ft)"] + bbox_tolerance_ft >= center_z)
        ]
        if not generic_matches.empty:
            nested_storefront_ids.add(str(storefront["ElementId"]))

    host_walls = exterior_walls[
        ~exterior_walls["ElementId"].astype(str).isin(nested_storefront_ids)
    ].copy()
    host_walls = host_walls.sort_values(
        ["Level", "element_center_y_ft", "element_center_x_ft", "ElementId"]
    ).reset_index(drop=True)

    attached_candidates = enriched[
        enriched["Category"].astype(str).isin(ATTACHED_CATEGORIES)
    ].copy()
    nested_storefronts = storefront_walls[
        storefront_walls["ElementId"].astype(str).isin(nested_storefront_ids)
    ].copy()
    attached_candidates = pd.concat([attached_candidates, nested_storefronts], ignore_index=True, sort=False)
    attached_candidates = attached_candidates.dropna(
        subset=["element_center_x_ft", "element_center_y_ft", "element_center_z_ft"]
    )

    prefab_rows: list[dict[str, object]] = []
    assignments: dict[str, tuple[str, float]] = {}
    host_by_group_id: dict[str, pd.Series] = {}
    candidate_by_id: dict[str, pd.Series] = {}

    for host_index, (_, wall) in enumerate(host_walls.iterrows(), start=1):
        wall_level = str(wall.get("Level", "")).strip()
        prefab_group_id = f"PREFAB_WALL_{normalized_level_token(wall_level)}_{host_index:03d}"
        wall_center = np.array(
            [
                float(wall["element_center_x_ft"]),
                float(wall["element_center_y_ft"]),
                float(wall["element_center_z_ft"]),
            ]
        )
        host_wall_element_id = str(wall["ElementId"])

        prefab_rows.append(
            {
                "prefab_group_id": prefab_group_id,
                "host_wall_element_id": host_wall_element_id,
                "element_id": host_wall_element_id,
                "category": wall.get("Category", ""),
                "family": wall.get("Family", ""),
                "type": wall.get("Type", ""),
                "level": wall_level,
                "takt_id": wall.get("takt_id", ""),
                "source_schedule": wall.get("source_schedule", ""),
            }
        )
        host_by_group_id[prefab_group_id] = wall

        min_x = float(wall["Bounding Box Min X (ft)"]) - bbox_tolerance_ft
        min_y = float(wall["Bounding Box Min Y (ft)"]) - bbox_tolerance_ft
        min_z = float(wall["Bounding Box Min Z (ft)"]) - bbox_tolerance_ft
        max_x = float(wall["Bounding Box Max X (ft)"]) + bbox_tolerance_ft
        max_y = float(wall["Bounding Box Max Y (ft)"]) + bbox_tolerance_ft
        max_z = float(wall["Bounding Box Max Z (ft)"]) + bbox_tolerance_ft

        nearby = attached_candidates[
            (attached_candidates["Level"].fillna("").astype(str).str.strip() == wall_level)
            & (attached_candidates["element_center_x_ft"] >= min_x)
            & (attached_candidates["element_center_x_ft"] <= max_x)
            & (attached_candidates["element_center_y_ft"] >= min_y)
            & (attached_candidates["element_center_y_ft"] <= max_y)
            & (attached_candidates["element_center_z_ft"] >= min_z)
            & (attached_candidates["element_center_z_ft"] <= max_z)
        ]

        for _, candidate in nearby.iterrows():
            candidate_id = str(candidate["ElementId"])
            candidate_by_id[candidate_id] = candidate
            candidate_center = np.array(
                [
                    float(candidate["element_center_x_ft"]),
                    float(candidate["element_center_y_ft"]),
                    float(candidate["element_center_z_ft"]),
                ]
            )
            distance = float(np.linalg.norm(candidate_center - wall_center))
            previous = assignments.get(candidate_id)
            if previous is not None and previous[1] <= distance:
                continue
            assignments[candidate_id] = (prefab_group_id, distance)

    for candidate_id, (prefab_group_id, _) in assignments.items():
        candidate = candidate_by_id[candidate_id]
        host_wall = host_by_group_id[prefab_group_id]
        prefab_rows.append(
            {
                "prefab_group_id": prefab_group_id,
                "host_wall_element_id": str(host_wall["ElementId"]),
                "element_id": candidate_id,
                "category": candidate.get("Category", ""),
                "family": candidate.get("Family", ""),
                "type": candidate.get("Type", ""),
                "level": candidate.get("Level", ""),
                "takt_id": candidate.get("takt_id", ""),
                "source_schedule": candidate.get("source_schedule", ""),
            }
        )

    prefab_df = pd.DataFrame(prefab_rows)
    if prefab_df.empty:
        return prefab_df

    prefab_df = prefab_df.drop_duplicates(subset=["prefab_group_id", "element_id"]).reset_index(drop=True)
    return prefab_df[
        [
            "prefab_group_id",
            "host_wall_element_id",
            "element_id",
            "category",
            "family",
            "type",
            "level",
            "takt_id",
            "source_schedule",
        ]
    ]


def main() -> None:
    OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)
    bim_df = load_central_bim()
    prefab_df = build_prefab_wall_mapping(bim_df)
    prefab_df.to_csv(PREFAB_MAPPING_OUTPUT_PATH, index=False)
    print(f"Wrote {PREFAB_MAPPING_OUTPUT_PATH.name}")
    print(f"Prefab groups: {prefab_df['prefab_group_id'].nunique() if not prefab_df.empty else 0}")
    print(f"Mapped rows: {len(prefab_df)}")


if __name__ == "__main__":
    main()
