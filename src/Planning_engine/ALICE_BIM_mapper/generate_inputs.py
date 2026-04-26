from __future__ import annotations

from pathlib import Path
import re

import pandas as pd


BASE_DIR = Path(__file__).resolve().parent
INPUTS_DIR = BASE_DIR / "inputs"
OUTPUTS_DIR = BASE_DIR / "outputs"

WORKBOOK_PATH = INPUTS_DIR / "ALICE_macro.xlsx"

MACRO_SCHEDULE_PATH = OUTPUTS_DIR / "Macro_Schedule.csv"
CREW_PATH = OUTPUTS_DIR / "Crew.csv"
EQUIPMENT_PATH = OUTPUTS_DIR / "Equipment.csv"
TASKS_PATH = OUTPUTS_DIR / "Tasks.csv"
MISSING_DATA_PATH = OUTPUTS_DIR / "Missing_data.md"

INCLUDED_WBS_NAMES = {
    "SITE PREPARATION & DEMOLITION",
    "EARTHWORK & BASEMENT",
    "SUPERSTRUCTURE",
    "ENVELOPE",
}


def slugify(value: str) -> str:
    text = re.sub(r"[^a-z0-9]+", "_", str(value).strip().lower())
    return text.strip("_")


def normalize_crew_type(name: str) -> str:
    return f"{slugify(name)}_crew" if not slugify(name).endswith("_crew") else slugify(name)


def normalize_equipment_type(name: str) -> str:
    return slugify(name)


def to_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def to_count_text(value: object) -> str:
    if pd.isna(value):
        return ""
    number = float(value)
    return str(int(number)) if number.is_integer() else str(number)


def join_unique(values: list[str]) -> str:
    cleaned = [value for value in values if value]
    return "|".join(dict.fromkeys(cleaned))


def calendar_hours(calendar_name: str) -> str:
    name = str(calendar_name).strip().lower()
    mapping = {
        "default calendar": "9,10,11,12,13,14,15,16,17",
        "5dx10hr": "8,9,10,11,12,13,14,15,16,17,18",
    }
    return mapping.get(name, "")


def load_included_wbs_ids(wbs: pd.DataFrame) -> set[str]:
    included = wbs[
        wbs["Name*"].astype(str).str.strip().str.upper().isin({name.upper() for name in INCLUDED_WBS_NAMES})
    ].copy()
    return set(included["Alice WBS Id*"].astype(str).str.strip())


def build_outputs() -> dict[str, list[str]]:
    workbook = pd.ExcelFile(WORKBOOK_PATH)

    wbs = workbook.parse("WBS")
    tasks = workbook.parse("Tasks")
    crews = workbook.parse("Crews")
    equipment = workbook.parse("Equipment")
    task_crews = workbook.parse("Task Crews")
    task_equipment = workbook.parse("Task Equipment")

    included_wbs_ids = load_included_wbs_ids(wbs)
    scoped_tasks = tasks[tasks["Alice WBS Id*"].astype(str).isin(included_wbs_ids)].copy()
    scoped_task_ids = set(scoped_tasks["Id*"].astype(str))

    task_crews = task_crews[task_crews["Task Id*"].astype(str).isin(scoped_task_ids)].copy()
    task_equipment = task_equipment[task_equipment["Task Id*"].astype(str).isin(scoped_task_ids)].copy()

    crew_lookup = {
        to_text(row["Name*"]): {
            "crew_type": normalize_crew_type(to_text(row["Name*"])),
            "count": to_count_text(row.get("Available quantity")),
            "cost": to_text(row.get("Cost per crew per hr*")),
            "hours": calendar_hours(to_text(row.get("Calendar*"))),
        }
        for _, row in crews.iterrows()
    }
    equipment_lookup = {
        to_text(row["Name*"]): {
            "equipment_type": normalize_equipment_type(to_text(row["Name*"])),
            "count": to_count_text(row.get("Available quantity")),
            "cost": to_text(row.get("Cost per hr*")),
        }
        for _, row in equipment.iterrows()
    }

    missing: dict[str, list[str]] = {"resource_lookup": [], "task_fields": []}

    macro_rows: list[dict[str, str]] = []
    task_rows: list[dict[str, str]] = []
    used_crews: set[str] = set()
    used_equipment: set[str] = set()

    for _, task_row in scoped_tasks.sort_values("Id*").iterrows():
        task_id = to_text(task_row["Id*"])
        task_name = to_text(task_row["Name*"])

        task_crew_rows = task_crews[task_crews["Task Id*"] == task_id]
        task_equipment_rows = task_equipment[task_equipment["Task Id*"] == task_id]

        crew_types: list[str] = []
        crew_counts: list[str] = []
        for _, crew_row in task_crew_rows.iterrows():
            crew_name = to_text(crew_row["Alice Crew Name*"])
            normalized = crew_lookup.get(crew_name, {}).get("crew_type", normalize_crew_type(crew_name))
            crew_types.append(normalized)
            crew_counts.append(to_count_text(crew_row.get("Required Amount")))
            used_crews.add(crew_name)
            if crew_name not in crew_lookup:
                missing["resource_lookup"].append(f"Crew lookup missing for `{crew_name}` used by `{task_id} {task_name}`.")

        equipment_types: list[str] = []
        for _, equipment_row in task_equipment_rows.iterrows():
            equipment_name = to_text(equipment_row["Alice Equipment Name*"])
            normalized = equipment_lookup.get(equipment_name, {}).get(
                "equipment_type", normalize_equipment_type(equipment_name)
            )
            equipment_types.append(normalized)
            used_equipment.add(equipment_name)
            if equipment_name not in equipment_lookup:
                missing["resource_lookup"].append(
                    f"Equipment lookup missing for `{equipment_name}` used by `{task_id} {task_name}`."
                )

        macro_rows.append(
            {
                "task_id": task_id,
                "task_name": task_name,
                "start_date": to_text(task_row["Planned Start Date - read only"]),
                "end_date": to_text(task_row["Planned End Date - read only"]),
                "crew_type": join_unique(crew_types),
                "equipment_type": join_unique(equipment_types),
            }
        )
        task_rows.append(
            {
                "task_name": task_name,
                "crew_type": join_unique(crew_types),
                "crew_num_req": join_unique(crew_counts),
                "equipment_type": join_unique(equipment_types),
            }
        )

    crew_rows: list[dict[str, str]] = []
    for crew_name in sorted(used_crews):
        info = crew_lookup.get(crew_name)
        if info is None:
            crew_rows.append({"crew_type": normalize_crew_type(crew_name), "count": "", "cost": "", "hours": ""})
            continue
        crew_rows.append(
            {
                "crew_type": info["crew_type"],
                "count": info["count"],
                "cost": info["cost"],
                "hours": info["hours"],
            }
        )

    equipment_rows: list[dict[str, str]] = []
    for equipment_name in sorted(used_equipment):
        info = equipment_lookup.get(equipment_name)
        if info is None:
            equipment_rows.append({"equipment_type": normalize_equipment_type(equipment_name), "count": "", "cost": ""})
            continue
        equipment_rows.append(
            {
                "equipment_type": info["equipment_type"],
                "count": info["count"],
                "cost": info["cost"],
            }
        )

    pd.DataFrame(
        macro_rows,
        columns=["task_id", "task_name", "start_date", "end_date", "crew_type", "equipment_type"],
    ).to_csv(MACRO_SCHEDULE_PATH, index=False)

    pd.DataFrame(crew_rows, columns=["crew_type", "count", "cost", "hours"]).to_csv(CREW_PATH, index=False)

    pd.DataFrame(equipment_rows, columns=["equipment_type", "count", "cost"]).to_csv(
        EQUIPMENT_PATH, index=False
    )

    pd.DataFrame(
        task_rows,
        columns=[
            "task_name",
            "crew_type",
            "crew_num_req",
            "equipment_type",
        ],
    ).to_csv(TASKS_PATH, index=False)

    return missing


def write_missing_data_report(missing: dict[str, list[str]]) -> None:
    lines = [
        "# Superstructure Missing Data",
        "",
        "These CSVs were generated from `ALICE_macro.xlsx` using the workbook WBS groups for site preparation, earthwork and basement, superstructure, and envelope.",
        "",
        "## Missing task-level data left blank",
        "",
    ]

    task_missing = sorted(dict.fromkeys(missing["task_fields"]))
    if task_missing:
        lines.extend(f"- {item}" for item in task_missing)
    else:
        lines.append("- None.")

    lines.extend(["", "## Missing resource lookups", ""])
    resource_missing = sorted(dict.fromkeys(missing["resource_lookup"]))
    if resource_missing:
        lines.extend(f"- {item}" for item in resource_missing)
    else:
        lines.append("- None.")

    lines.extend(
        [
        ]
    )

    MISSING_DATA_PATH.write_text("\n".join(lines) + "\n", encoding="utf-8")


if __name__ == "__main__":
    missing_data = build_outputs()
    write_missing_data_report(missing_data)
    print(f"Wrote {MACRO_SCHEDULE_PATH.name}")
    print(f"Wrote {CREW_PATH.name}")
    print(f"Wrote {EQUIPMENT_PATH.name}")
    print(f"Wrote {TASKS_PATH.name}")
    print(f"Wrote {MISSING_DATA_PATH.name}")
