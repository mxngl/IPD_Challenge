from __future__ import annotations

import json
from pathlib import Path

import matplotlib.image as mpimg
import matplotlib.pyplot as plt
from matplotlib.patches import Polygon
from matplotlib.path import Path as MplPath
import numpy as np
import pandas as pd


PROJECT_DIR = Path(__file__).resolve().parent.parent
OUTPUT_PATH = PROJECT_DIR / "outputs" / "takt_zones" / "takt_zones.json"
CENTRAL_BIM_OUTPUT_PATH = PROJECT_DIR / "outputs" / "takt_zones" / "central_bim_model.csv"
CENTRAL_BIM_WITH_TAKT_OUTPUT_PATH = PROJECT_DIR / "outputs" / "takt_zones" / "central_bim_model_with_takt.csv"

BACKGROUND_BY_LEVEL = {
    "L -1": PROJECT_DIR / "floor_plans" / "L-1.png",
    "L 0": PROJECT_DIR / "floor_plans" / "L0.png",
    "L 1": PROJECT_DIR / "floor_plans" / "L1.png",
}

ARCHITECTURE_SCHEDULE_PATH = (
    PROJECT_DIR
    / "revit_schedules"
    / "Arch"
    / "04_Island_ARCH_Concept2_V24_2026-04-23_04-35-08pm_detached_Architecture_TakeOff.csv"
)
STRUCT_ARCHITECTURE_SCHEDULE_PATH = (
    PROJECT_DIR
    / "revit_schedules"
    / "Struct"
    / "STR_Wall_Bamboo_Concept2_amd03_V3_2026-04-24_08-20-35am_detached_Architecture_TakeOff.csv"
)
ARCHITECTURE_STRUCTURAL_SCHEDULE_PATH = (
    PROJECT_DIR
    / "revit_schedules"
    / "Arch"
    / "04_Island_ARCH_Concept2_V24_2026-04-23_04-35-08pm_detached_Structural_Schedule.csv"
)
STRUCT_STRUCTURAL_SCHEDULE_PATH = (
    PROJECT_DIR
    / "revit_schedules"
    / "Struct"
    / "STR_Wall_Bamboo_Concept2_amd03_V3_2026-04-24_08-20-35am_detached_Structural_Schedule.csv"
)
ARCHITECTURE_MEP_SCHEDULE_PATH = (
    PROJECT_DIR
    / "revit_schedules"
    / "Arch"
    / "04_Island_ARCH_Concept2_V24_2026-04-23_04-35-08pm_detached_MEP_TakeOff.csv"
)

MEP_SCHEDULE_PATH = (
PROJECT_DIR
    / "revit_schedules"
    / "MEP"
    / "01_Island_MEP_Concept2_V4_2026-04-08_11-47-28am_detached_MEP_TakeOff.csv"
)

LEVEL_SEQUENCE = ["L -1", "L 0", "L 1"]
ZONE_FACE_COLORS = ["#ff6b6b", "#4dabf7", "#51cf66", "#ffd43b", "#b197fc", "#ffa94d"]
BIM_SCHEDULE_PATHS = [
    ARCHITECTURE_SCHEDULE_PATH,
    STRUCT_ARCHITECTURE_SCHEDULE_PATH,
    ARCHITECTURE_STRUCTURAL_SCHEDULE_PATH,
    STRUCT_STRUCTURAL_SCHEDULE_PATH,
    ARCHITECTURE_MEP_SCHEDULE_PATH,
]


def load_plot_elements(schedule_paths: list[Path]) -> pd.DataFrame:
    frames: list[pd.DataFrame] = []
    coord_columns = [
        "Position X (ft)",
        "Position Y (ft)",
        "Position Z (ft)",
        "Bounding Box Min X (ft)",
        "Bounding Box Min Y (ft)",
        "Bounding Box Min Z (ft)",
        "Bounding Box Max X (ft)",
        "Bounding Box Max Y (ft)",
        "Bounding Box Max Z (ft)",
    ]
    included_categories = {"Walls", "Curtain Wall Mullions", "Curtain Panels"}

    for schedule_path in schedule_paths:
        df = pd.read_csv(schedule_path)
        plot_elements = df[df["Category"].astype(str).isin(included_categories)].copy()
        plot_elements["source_schedule"] = str(schedule_path)
        for column in coord_columns:
            plot_elements[column] = pd.to_numeric(plot_elements[column], errors="coerce")
        frames.append(plot_elements.dropna(subset=coord_columns))

    if not frames:
        return pd.DataFrame(columns=coord_columns + ["source_schedule"])

    combined = pd.concat(frames, ignore_index=True)
    dedupe_columns = [
        "Level",
        "Category",
        "Type",
        "Position X (ft)",
        "Position Y (ft)",
        "Position Z (ft)",
        "Bounding Box Min X (ft)",
        "Bounding Box Min Y (ft)",
        "Bounding Box Min Z (ft)",
        "Bounding Box Max X (ft)",
        "Bounding Box Max Y (ft)",
        "Bounding Box Max Z (ft)",
    ]
    return combined.drop_duplicates(subset=dedupe_columns).reset_index(drop=True)


def load_and_combine_bim_schedules(schedule_paths: list[Path]) -> pd.DataFrame:
    frames: list[pd.DataFrame] = []
    numeric_columns = [
        "Position X (ft)",
        "Position Y (ft)",
        "Position Z (ft)",
        "Bounding Box Min X (ft)",
        "Bounding Box Min Y (ft)",
        "Bounding Box Min Z (ft)",
        "Bounding Box Max X (ft)",
        "Bounding Box Max Y (ft)",
        "Bounding Box Max Z (ft)",
        "Bounding Box Center X (ft)",
        "Bounding Box Center Y (ft)",
        "Bounding Box Center Z (ft)",
    ]

    for schedule_path in schedule_paths:
        if not schedule_path.exists():
            print(f"Skipping missing BIM schedule: {schedule_path}")
            continue
        df = pd.read_csv(schedule_path)
        df["source_schedule"] = str(schedule_path)
        for column in numeric_columns:
            if column in df.columns:
                df[column] = pd.to_numeric(df[column], errors="coerce")
        frames.append(df)

    if not frames:
        return pd.DataFrame()

    combined = pd.concat(frames, ignore_index=True, sort=False)
    dedupe_columns = [
        column
        for column in [
            "Category",
            "Family",
            "Type",
            "Level",
            "Position X (ft)",
            "Position Y (ft)",
            "Position Z (ft)",
            "Bounding Box Min X (ft)",
            "Bounding Box Min Y (ft)",
            "Bounding Box Min Z (ft)",
            "Bounding Box Max X (ft)",
            "Bounding Box Max Y (ft)",
            "Bounding Box Max Z (ft)",
        ]
        if column in combined.columns
    ]
    combined = combined.drop_duplicates(subset=dedupe_columns).reset_index(drop=True)
    return combined


def get_element_center(row: pd.Series) -> tuple[float | None, float | None]:
    center_x = row.get("Bounding Box Center X (ft)")
    center_y = row.get("Bounding Box Center Y (ft)")
    if pd.notna(center_x) and pd.notna(center_y):
        return float(center_x), float(center_y)

    position_x = row.get("Position X (ft)")
    position_y = row.get("Position Y (ft)")
    if pd.notna(position_x) and pd.notna(position_y):
        return float(position_x), float(position_y)

    return None, None


def assign_takt_ids(bim_df: pd.DataFrame, zones_by_level: dict[str, list[dict[str, object]]]) -> pd.DataFrame:
    enriched = bim_df.copy()
    enriched["takt_id"] = ""

    for row_index, row in enriched.iterrows():
        level = str(row.get("Level", "")).strip()
        level_zones = zones_by_level.get(level, [])
        if not level_zones:
            continue

        center_x, center_y = get_element_center(row)
        if center_x is None or center_y is None:
            continue

        point = (center_x, center_y)
        for zone in level_zones:
            polygon_points = np.array(zone["corners_model_xy"], dtype=float)
            polygon_path = MplPath(polygon_points, closed=True)
            if polygon_path.contains_point(point, radius=1e-9):
                enriched.at[row_index, "takt_id"] = zone["zone_name"]
                break

    return enriched


def wall_centerline(row: pd.Series) -> np.ndarray:
    span_x = row["Bounding Box Max X (ft)"] - row["Bounding Box Min X (ft)"]
    span_y = row["Bounding Box Max Y (ft)"] - row["Bounding Box Min Y (ft)"]
    if span_x >= span_y:
        return np.array(
            [
                [row["Bounding Box Min X (ft)"], row["Position Y (ft)"]],
                [row["Bounding Box Max X (ft)"], row["Position Y (ft)"]],
            ]
        )
    return np.array(
        [
            [row["Position X (ft)"], row["Bounding Box Min Y (ft)"]],
            [row["Position X (ft)"], row["Bounding Box Max Y (ft)"]],
        ]
    )


def build_model_bounds(elements_2d: pd.DataFrame, padding: float = 8.0) -> dict[str, float]:
    all_points = np.vstack([wall_centerline(row) for _, row in elements_2d.iterrows()])
    return {
        "min_x": float(all_points[:, 0].min() - padding),
        "max_x": float(all_points[:, 0].max() + padding),
        "min_y": float(all_points[:, 1].min() - padding),
        "max_y": float(all_points[:, 1].max() + padding),
    }


def transform_points(points: np.ndarray, affine: np.ndarray) -> np.ndarray:
    homogeneous = np.column_stack([points, np.ones(len(points))])
    return homogeneous @ affine.T


def inverse_transform_points(points: np.ndarray, affine: np.ndarray) -> np.ndarray:
    affine_augmented = np.vstack([affine, np.array([0.0, 0.0, 1.0])])
    inverse_affine = np.linalg.inv(affine_augmented)
    homogeneous = np.column_stack([points, np.ones(len(points))])
    return (homogeneous @ inverse_affine.T)[:, :2]


def calibrate_level(level: str, elements_2d: pd.DataFrame, image: np.ndarray) -> np.ndarray:
    model_bounds = build_model_bounds(elements_2d)
    fig, axes = plt.subplots(1, 2, figsize=(16, 8))
    model_ax, image_ax = axes

    for _, row in elements_2d.iterrows():
        centerline = wall_centerline(row)
        model_ax.plot(centerline[:, 0], centerline[:, 1], color="deepskyblue", linewidth=2.0, alpha=0.7)
        model_ax.scatter(row["Position X (ft)"], row["Position Y (ft)"], color="darkorange", s=12, alpha=0.7)

    model_ax.set_title(f"{level}: click 4 matching reference points in model space")
    model_ax.set_xlim(model_bounds["min_x"], model_bounds["max_x"])
    model_ax.set_ylim(model_bounds["min_y"], model_bounds["max_y"])
    model_ax.set_aspect("equal")
    model_ax.grid(alpha=0.18)
    model_ax.set_xlabel("Model X (ft)")
    model_ax.set_ylabel("Model Y (ft)")

    image_ax.imshow(image)
    image_ax.set_title(f"{level}: click the same 4 points on the cropped PNG")
    image_ax.axis("off")

    plt.tight_layout()
    plt.show(block=False)

    print(f"{level}: click 4 reference points on the LEFT model plot in order, then press Enter.")
    model_clicks = np.array(plt.ginput(4, timeout=-1), dtype=float)
    if len(model_clicks) != 4:
        plt.close(fig)
        raise ValueError(f"{level}: expected 4 model-space clicks.")
    model_ax.scatter(model_clicks[:, 0], model_clicks[:, 1], color="crimson", s=70, marker="x")
    fig.canvas.draw_idle()

    print(f"{level}: click the same 4 reference points on the RIGHT cropped PNG in the same order, then press Enter.")
    image_clicks = np.array(plt.ginput(4, timeout=-1), dtype=float)
    if len(image_clicks) != 4:
        plt.close(fig)
        raise ValueError(f"{level}: expected 4 image-space clicks.")
    image_ax.scatter(image_clicks[:, 0], image_clicks[:, 1], color="crimson", s=70, marker="x")
    fig.canvas.draw_idle()
    plt.close(fig)

    design = np.column_stack([model_clicks, np.ones(len(model_clicks))])
    affine_x, *_ = np.linalg.lstsq(design, image_clicks[:, 0], rcond=None)
    affine_y, *_ = np.linalg.lstsq(design, image_clicks[:, 1], rcond=None)
    return np.vstack([affine_x, affine_y])


class ZoneCollector:
    def __init__(self, figure, axis, level: str, elements_2d: pd.DataFrame, image: np.ndarray, affine: np.ndarray):
        self.figure = figure
        self.ax = axis
        self.level = level
        self.elements_2d = elements_2d
        self.image = image
        self.affine = affine
        self.current_points: list[list[float]] = []
        self.completed_zones: list[dict[str, object]] = []
        self.finished_floor = False
        self.preview_artist = None

        self.figure.canvas.mpl_connect("button_press_event", self.on_click)
        self.figure.canvas.mpl_connect("key_press_event", self.on_key)

        self.draw_background()

    def draw_background(self) -> None:
        self.ax.clear()
        self.ax.imshow(self.image)

        for _, row in self.elements_2d.iterrows():
            centerline = wall_centerline(row)
            transformed_line = transform_points(centerline, self.affine)
            transformed_center = transform_points(
                np.array([[row["Position X (ft)"], row["Position Y (ft)"]]], dtype=float),
                self.affine,
            )[0]
            self.ax.plot(
                transformed_line[:, 0],
                transformed_line[:, 1],
                color="deepskyblue",
                linewidth=1.8,
                alpha=0.55,
            )
            self.ax.scatter(
                transformed_center[0],
                transformed_center[1],
                color="darkorange",
                s=12,
                alpha=0.7,
                edgecolors="white",
                linewidth=0.35,
            )

        self.ax.set_title(
            f"{self.level}: click takt zone corners on the cropped PNG. "
            "Press o = save zone, Esc = next floor"
        )
        self.ax.axis("off")

        for zone in self.completed_zones:
            self.add_zone_artists(
                np.array(zone["corners_image_xy"], dtype=float),
                zone["zone_name"],
                zone["face_color"],
            )

        self.figure.canvas.draw_idle()

    def add_zone_artists(
        self,
        image_points: np.ndarray,
        zone_name: str,
        face_color: str,
        alpha: float = 0.25,
    ) -> None:
        patch = Polygon(
            image_points,
            closed=True,
            facecolor=face_color,
            edgecolor=face_color,
            alpha=alpha,
            linewidth=2,
        )
        self.ax.add_patch(patch)
        label_xy = image_points.mean(axis=0)
        self.ax.text(
            label_xy[0],
            label_xy[1],
            zone_name,
            color="#1f2937",
            fontsize=9,
            ha="center",
            va="center",
            bbox={"facecolor": "white", "edgecolor": "none", "alpha": 0.75, "pad": 1.5},
        )

    def draw_current_zone(self) -> None:
        if self.preview_artist is not None:
            self.preview_artist.remove()
            self.preview_artist = None

        if len(self.current_points) >= 2:
            preview_points = np.array(self.current_points, dtype=float)
            self.preview_artist = Polygon(
                preview_points,
                closed=False,
                fill=False,
                edgecolor="#212529",
                linestyle="--",
                linewidth=1.5,
            )
            self.ax.add_patch(self.preview_artist)

        self.figure.canvas.draw_idle()

    def on_click(self, event) -> None:
        if event.inaxes is not self.ax or self.finished_floor:
            return
        if event.xdata is None or event.ydata is None:
            return
        self.current_points.append([float(event.xdata), float(event.ydata)])
        self.ax.scatter(event.xdata, event.ydata, color="crimson", s=55, marker="x")
        self.draw_current_zone()

    def on_key(self, event) -> None:
        if self.finished_floor:
            return

        if event.key == "o":
            if len(self.current_points) < 3:
                print(f"{self.level}: current zone has {len(self.current_points)} points. Click at least 3 corners first.")
                return
            self.finalize_zone()
        elif event.key == "escape":
            if self.current_points:
                print(f"{self.level}: discarding {len(self.current_points)} in-progress points and finishing the floor.")
                self.current_points = []
                self.draw_background()
            self.finished_floor = True
            plt.close(self.figure)

    def finalize_zone(self) -> None:
        image_points = np.array(self.current_points, dtype=float)
        model_points = inverse_transform_points(image_points, self.affine)
        zone_index = len(self.completed_zones) + 1
        face_color = ZONE_FACE_COLORS[(zone_index - 1) % len(ZONE_FACE_COLORS)]
        zone_name = f"{self.level} Zone {zone_index}"

        zone_payload = {
            "zone_name": zone_name,
            "level": self.level,
            "corners_image_xy": np.round(image_points, 3).tolist(),
            "corners_model_xy": np.round(model_points, 3).tolist(),
            "face_color": face_color,
        }
        self.completed_zones.append(zone_payload)
        self.current_points = []
        self.draw_background()
        print(
            f"{self.level}: saved {zone_name}. "
            "Keep clicking corners on the PNG and press o to save another zone, or Esc for next floor."
        )


def collect_zones_for_level(level: str, elements_2d: pd.DataFrame, image: np.ndarray, affine: np.ndarray) -> list[dict[str, object]]:
    fig, ax = plt.subplots(figsize=(11, 11))
    collector = ZoneCollector(fig, ax, level, elements_2d, image, affine)
    plt.tight_layout()
    plt.show()
    return collector.completed_zones


def save_zones(schedule_paths: list[Path], zone_data: dict[str, object]) -> None:
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_PATH.write_text(json.dumps(zone_data, indent=2), encoding="utf-8")
    print(f"Saved takt zones to: {OUTPUT_PATH}")


def save_central_bim_exports(bim_df: pd.DataFrame, enriched_bim_df: pd.DataFrame) -> None:
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    bim_df.to_csv(CENTRAL_BIM_OUTPUT_PATH, index=False)
    enriched_bim_df.to_csv(CENTRAL_BIM_WITH_TAKT_OUTPUT_PATH, index=False)
    print(f"Saved combined BIM model to: {CENTRAL_BIM_OUTPUT_PATH}")
    print(f"Saved combined BIM model with takt ids to: {CENTRAL_BIM_WITH_TAKT_OUTPUT_PATH}")


def load_existing_takt_zones() -> dict[str, object] | None:
    if not OUTPUT_PATH.exists():
        return None
    return json.loads(OUTPUT_PATH.read_text(encoding="utf-8"))


def main() -> None:
    plot_schedule_paths = [ARCHITECTURE_SCHEDULE_PATH, STRUCT_ARCHITECTURE_SCHEDULE_PATH]
    all_elements = load_plot_elements(plot_schedule_paths)
    existing_zone_data = load_existing_takt_zones()

    if existing_zone_data is None:
        zones_by_level: dict[str, list[dict[str, object]]] = {}
        calibration_affine_by_level: dict[str, list[list[float]]] = {}

        for level in LEVEL_SEQUENCE:
            elements_2d = all_elements[all_elements["Level"].astype(str).eq(level)].copy()
            if elements_2d.empty:
                print(f"Skipping {level}: no plotted elements found.")
                continue

            image_path = BACKGROUND_BY_LEVEL.get(level)
            if image_path is None or not image_path.exists():
                print(f"Skipping {level}: missing cropped PNG.")
                continue

            image = mpimg.imread(image_path)
            affine = calibrate_level(level, elements_2d, image)
            calibration_affine_by_level[level] = np.round(affine, 6).tolist()

            print(
                f"{level}: click takt zone corners on the cropped PNG in order. "
                "Press o to save the current polygon, or Esc to finish this floor."
            )
            zones_by_level[level] = collect_zones_for_level(level, elements_2d, image, affine)

        existing_zone_data = {
            "schedule_paths": [str(path) for path in BIM_SCHEDULE_PATHS],
            "floor_plan_image_by_level": {level: str(path) for level, path in BACKGROUND_BY_LEVEL.items()},
            "calibration_affine_by_level": calibration_affine_by_level,
            "levels": zones_by_level,
        }
        save_zones(BIM_SCHEDULE_PATHS, existing_zone_data)
    else:
        print(f"Using existing takt zones from: {OUTPUT_PATH}")

    combined_bim_df = load_and_combine_bim_schedules(BIM_SCHEDULE_PATHS)
    if combined_bim_df.empty:
        print("No BIM schedules were combined, so takt ids were not assigned.")
        return

    zones_by_level = existing_zone_data.get("levels", {})
    enriched_bim_df = assign_takt_ids(combined_bim_df, zones_by_level)
    save_central_bim_exports(combined_bim_df, enriched_bim_df)


if __name__ == "__main__":
    main()
