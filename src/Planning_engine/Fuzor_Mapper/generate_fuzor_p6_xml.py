from __future__ import annotations

import json
from pathlib import Path
import uuid
import xml.etree.ElementTree as ET

import pandas as pd


BASE_DIR = Path(__file__).resolve().parent
INPUTS_DIR = BASE_DIR / "inputs"
OUTPUTS_DIR = BASE_DIR / "outputs"
PLANNING_ENGINE_DIR = BASE_DIR.parent
MICRO_SCHEDULE_PATH = PLANNING_ENGINE_DIR / "Micro_Schedule_Generator" / "outputs" / "Micro_Schedule.csv"
OUTPUT_XML_PATH = OUTPUTS_DIR / "Fuzor_Micro_Schedule.xml"
REVIT_BUILD_CODE_MAP_PATH = OUTPUTS_DIR / "Revit_4D_Build_Code_Map.csv"

NS = "http://xmlns.oracle.com/Primavera/P6/V25.12/API/BusinessObjects"
XSI = "http://www.w3.org/2001/XMLSchema-instance"

ET.register_namespace("", NS)
ET.register_namespace("xsi", XSI)


def qname(tag: str) -> str:
    return f"{{{NS}}}{tag}"


def add_text(parent: ET.Element, tag: str, value: object) -> ET.Element:
    element = ET.SubElement(parent, qname(tag))
    element.text = "" if value is None else str(value)
    return element


def p6_guid() -> str:
    return "{" + str(uuid.uuid4()).upper() + "}"


def format_dt(value: object) -> str:
    ts = pd.to_datetime(value)
    return ts.isoformat(timespec="seconds")


def safe_activity_id(task_id: str, level: str, element_id: str) -> str:
    level_token = (
        str(level)
        .replace(" ", "")
        .replace("-", "M")
        .replace("+", "P")
        .replace("/", "_")
    )
    return f"{task_id}_{level_token}_{element_id}"


def normalize_level(value: object, task_name: object = "") -> str:
    if pd.isna(value) or str(value).strip() == "":
        if str(task_name).strip().lower() == "contingency":
            return "Project Contingency"
        return "Unassigned"
    return str(value)


def normalize_element_id(value: object, task_id: object) -> str:
    if pd.isna(value) or str(value).strip() == "":
        return str(task_id)
    text = str(value).strip()
    try:
        numeric = float(text)
    except ValueError:
        return text
    if numeric.is_integer():
        return str(int(numeric))
    return text


def level_sort_key(level: object) -> tuple[float, str]:
    text = str(level).strip()
    if not text:
        return (9999.0, text)
    if text.lower() == "basement / site":
        return (-100.0, text)
    normalized = text.lower()
    if normalized == "roof":
        return (9000.0, text)
    if normalized == "project contingency":
        return (9800.0, text)
    digits = "".join(ch if ch.isdigit() or ch == "-" else " " for ch in text).split()
    if digits:
        return (float(digits[0]), text)
    return (9500.0, text)


def build_activity_table() -> pd.DataFrame:
    df = pd.read_csv(MICRO_SCHEDULE_PATH)
    activities = (
        df.drop_duplicates(["task_id", "element_key", "element_start", "element_end"])
        .copy()
        .sort_values(["element_start", "task_id", "order_in_task", "element_key"])
        .reset_index(drop=True)
    )
    activities["level"] = activities.apply(
        lambda row: normalize_level(row["level"], row["task_name"]),
        axis=1,
    )
    activities["element_id"] = activities.apply(
        lambda row: normalize_element_id(row["element_id"], row["task_id"]),
        axis=1,
    )
    activities["level_sort"] = activities["level"].apply(level_sort_key)
    activities["activity_id"] = activities.apply(
        lambda row: safe_activity_id(str(row["task_id"]), str(row["level"]), str(row["element_id"])),
        axis=1,
    )
    activities["export_task_name"] = activities.apply(
        lambda row: f"{row['task_name']} | {row['level']} | {row['element_id']}",
        axis=1,
    )
    return activities


def build_xml() -> ET.ElementTree:
    activities = build_activity_table()

    project_start = format_dt(activities["element_start"].min())
    project_finish = format_dt(activities["element_end"].max())

    root = ET.Element(
        qname("APIBusinessObjects"),
        {f"{{{XSI}}}schemaLocation": ""},
    )

    currency = ET.SubElement(root, qname("Currency"))
    add_text(currency, "DecimalPlaces", 2)
    add_text(currency, "DecimalSymbol", "Period")
    add_text(currency, "DigitGroupingSymbol", "Comma")
    add_text(currency, "ExchangeRate", 1)
    add_text(currency, "Id", "CUR")
    add_text(currency, "Name", "Default Currency")
    add_text(currency, "NegativeSymbol", "(#1.1)")
    add_text(currency, "ObjectId", 1)
    add_text(currency, "PositiveSymbol", "#1.1")
    add_text(currency, "Symbol", "$")

    calendar = ET.SubElement(root, qname("Calendar"))
    add_text(calendar, "HoursPerDay", 8)
    add_text(calendar, "HoursPerMonth", 160)
    add_text(calendar, "HoursPerWeek", 40)
    add_text(calendar, "HoursPerYear", 1920)
    add_text(calendar, "IsDefault", 1)
    add_text(calendar, "IsPersonal", 0)
    add_text(calendar, "Name", "Default calendar")
    add_text(calendar, "ObjectId", 4)
    add_text(calendar, "Type", "Global")
    standard_work_week = ET.SubElement(calendar, qname("StandardWorkWeek"))
    for day in ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]:
        day_hours = ET.SubElement(standard_work_week, qname("StandardWorkHours"))
        add_text(day_hours, "DayOfWeek", day)
        if day not in {"Sunday", "Saturday"}:
            work_time = ET.SubElement(day_hours, qname("WorkTime"))
            add_text(work_time, "Start", "09:00:00")
            add_text(work_time, "Finish", "16:59:00")
    ET.SubElement(calendar, qname("HolidayOrExceptions"))

    project = ET.SubElement(root, qname("Project"))
    add_text(project, "ActivityDefaultActivityType", "Task Dependent")
    add_text(project, "ActivityDefaultCalendarObjectId", 4)
    add_text(project, "ActivityDefaultDurationType", "Fixed Duration and Units")
    add_text(project, "ActivityDefaultPercentCompleteType", "Duration")
    add_text(project, "ActivityDefaultPricePerUnit", 0)
    add_text(project, "ActivityIdBasedOnSelectedActivity", 1)
    add_text(project, "ActivityIdIncrement", 10)
    add_text(project, "ActivityIdPrefix", "M")
    add_text(project, "ActivityIdSuffix", "1000")
    add_text(project, "ActivityPercentCompleteBasedOnActivitySteps", 0)
    add_text(project, "AddActualToRemaining", 0)
    add_text(project, "AllowNegativeActualUnitsFlag", 0)
    add_text(project, "AssignmentDefaultDrivingFlag", 1)
    add_text(project, "AssignmentDefaultRateType", "Price / Unit")
    add_text(project, "CheckOutStatus", 0)
    add_text(project, "CostQuantityRecalculateFlag", 0)
    add_text(project, "CriticalActivityFloatLimit", 0)
    add_text(project, "CriticalActivityPathType", "Critical Float")
    add_text(project, "DataDate", project_start)
    add_text(project, "DefaultPriceTimeUnits", "Hour")
    add_text(project, "DiscountApplicationPeriod", "Month")
    add_text(project, "EarnedValueComputeType", "Activity Percent Complete")
    add_text(project, "EarnedValueETCComputeType", "ETC = Remaining Cost for Activity")
    add_text(project, "EarnedValueETCUserValue", 0.88)
    add_text(project, "EarnedValueUserPercent", 0.06)
    add_text(project, "EnablePublication", 0)
    add_text(project, "EnableSummarization", 0)
    add_text(project, "FinishDate", project_finish)
    add_text(project, "FiscalYearStartMonth", 1)
    add_text(project, "GUID", p6_guid())
    add_text(project, "Id", "Micro Basement + Superstructure")
    add_text(project, "LevelingPriority", 10)
    add_text(project, "LinkActualToActualThisPeriod", 1)
    add_text(project, "LinkPercentCompleteWithActual", 1)
    add_text(project, "LinkPlannedAndAtCompletionFlag", 1)
    add_text(project, "Name", "Micro Basement + Superstructure")
    add_text(project, "ObjectId", 1)
    add_text(project, "PlannedStartDate", project_start)
    add_text(project, "ProjectFlag", 1)
    add_text(project, "StartDate", project_start)
    add_text(project, "Status", "Active")

    project_calendar = ET.SubElement(project, qname("Calendar"))
    add_text(project_calendar, "ObjectId", 4)

    wbs_root_object_id = 100
    task_wbs_map: dict[tuple[str, str], int] = {}

    root_wbs = ET.SubElement(project, qname("WBS"))
    add_text(root_wbs, "Code", "1")
    add_text(root_wbs, "GUID", p6_guid())
    add_text(root_wbs, "Name", "BASEMENT + SUPERSTRUCTURE MICRO")
    add_text(root_wbs, "ObjectId", wbs_root_object_id)
    add_text(root_wbs, "ParentObjectId", "")
    add_text(root_wbs, "ProjectObjectId", 1)
    add_text(root_wbs, "SequenceNumber", 1)
    add_text(root_wbs, "Status", "Active")

    next_wbs_object_id = 110
    ordered_tasks = (
        activities.loc[:, ["task_id", "task_name", "element_start"]]
        .drop_duplicates(subset=["task_id", "task_name"])
        .sort_values(["element_start", "task_id", "task_name"])
        .reset_index(drop=True)
    )

    for task_index, row in enumerate(ordered_tasks.itertuples(index=False), start=1):
        task_key = (str(row.task_id), str(row.task_name))
        task_wbs_map[task_key] = next_wbs_object_id
        task_wbs = ET.SubElement(project, qname("WBS"))
        add_text(task_wbs, "Code", f"1.{task_index}")
        add_text(task_wbs, "GUID", p6_guid())
        add_text(task_wbs, "Name", str(row.task_name))
        add_text(task_wbs, "ObjectId", next_wbs_object_id)
        add_text(task_wbs, "ParentObjectId", wbs_root_object_id)
        add_text(task_wbs, "ProjectObjectId", 1)
        add_text(task_wbs, "SequenceNumber", task_index)
        add_text(task_wbs, "Status", "Active")
        next_wbs_object_id += 10

    activity_object_id_map: dict[str, int] = {}
    next_activity_object_id = 1000

    for _, row in activities.iterrows():
        activity_id = safe_activity_id(str(row["task_id"]), str(row["level"]), str(row["element_id"]))
        activity_object_id_map[activity_id] = next_activity_object_id

        activity = ET.SubElement(project, qname("Activity"))
        start_dt = format_dt(row["element_start"])
        finish_dt = format_dt(row["element_end"])
        duration_hours = round(float(row["scheduled_duration_hr"]), 4)
        primary_resource = 8 if str(row["productivity_dependency"]).strip().lower() == "equipment" else ""

        add_text(activity, "ActualLaborUnits", 0)
        add_text(activity, "ActualNonLaborUnits", 0)
        add_text(activity, "AtCompletionDuration", duration_hours)
        add_text(activity, "AutoComputeActuals", 1)
        add_text(activity, "CalendarObjectId", 4)
        add_text(activity, "DurationType", "Fixed Units/Time")
        add_text(activity, "FinishDate", finish_dt)
        add_text(activity, "GUID", p6_guid())
        add_text(activity, "Id", activity_id)
        add_text(activity, "LevelingPriority", "Normal")
        add_text(activity, "Name", str(row["export_task_name"]))
        add_text(activity, "ObjectId", next_activity_object_id)
        add_text(activity, "PercentCompleteType", "Duration")
        add_text(activity, "PlannedDuration", duration_hours)
        add_text(activity, "PlannedFinishDate", finish_dt)
        add_text(activity, "PlannedLaborUnits", duration_hours)
        add_text(activity, "PlannedNonLaborUnits", 0)
        add_text(activity, "PlannedStartDate", start_dt)
        add_text(activity, "PrimaryResourceObjectId", primary_resource)
        add_text(activity, "ProjectObjectId", 1)
        add_text(activity, "RemainingDuration", duration_hours)
        add_text(activity, "RemainingEarlyFinishDate", finish_dt)
        add_text(activity, "RemainingEarlyStartDate", start_dt)
        add_text(activity, "RemainingLaborCost", 0)
        add_text(activity, "RemainingLaborUnits", duration_hours)
        add_text(activity, "RemainingLateFinishDate", finish_dt)
        add_text(activity, "RemainingLateStartDate", start_dt)
        add_text(activity, "RemainingNonLaborCost", 0)
        add_text(activity, "RemainingNonLaborUnits", 0)
        add_text(activity, "StartDate", start_dt)
        add_text(activity, "Status", "Not Started")
        add_text(activity, "Type", "Task Dependent")
        add_text(activity, "WBSObjectId", task_wbs_map[(str(row["task_id"]), str(row["task_name"]))])

        note_public = ET.SubElement(project, qname("ActivityNote"))
        add_text(note_public, "ActivityObjectId", next_activity_object_id)
        add_text(
            note_public,
            "Note",
            (
                f"&lt;b&gt;Model Element:&lt;/b&gt; {row['element_id']}"
                f" &lt;br/&gt;&lt;b&gt;Level:&lt;/b&gt; {row['level']}"
                f" &lt;br/&gt;&lt;b&gt;Task:&lt;/b&gt; {row['task_name']}"
            ),
        )
        add_text(note_public, "NotebookTopicObjectId", 2)
        add_text(note_public, "ObjectId", next_activity_object_id + 1)
        add_text(note_public, "ProjectObjectId", 1)

        note_internal = ET.SubElement(project, qname("ActivityNote"))
        add_text(note_internal, "ActivityObjectId", next_activity_object_id)
        add_text(
            note_internal,
            "Note",
            json.dumps(
                {
                    "elementKey": row["element_key"],
                    "elementId": str(row["element_id"]),
                    "level": str(row["level"]),
                    "category": str(row["category"]),
                    "family": str(row["family"]),
                    "type": str(row["type"]),
                    "sourceModel": str(row["source_model"]),
                    "coordXft": float(row["coord_x_ft"]),
                    "coordYft": float(row["coord_y_ft"]),
                    "coordZft": float(row["coord_z_ft"]),
                }
            ),
        )
        add_text(note_internal, "NotebookTopicObjectId", 3)
        add_text(note_internal, "ObjectId", next_activity_object_id + 2)
        add_text(note_internal, "ProjectObjectId", 1)

        next_activity_object_id += 10

    relationship_object_id = 5000
    ordered_activities = activities.sort_values(
        ["element_start", "task_id", "order_in_task", "element_key"]
    ).reset_index(drop=True)
    previous_activity_id: str | None = None

    for _, row in ordered_activities.iterrows():
        current_activity_id = str(row["activity_id"])
        if previous_activity_id is not None:
            rel = ET.SubElement(project, qname("Relationship"))
            add_text(rel, "Lag", 0)
            add_text(rel, "ObjectId", relationship_object_id)
            add_text(rel, "PredecessorActivityObjectId", activity_object_id_map[previous_activity_id])
            add_text(rel, "PredecessorProjectObjectId", 1)
            add_text(rel, "SuccessorActivityObjectId", activity_object_id_map[current_activity_id])
            add_text(rel, "SuccessorProjectObjectId", 1)
            add_text(rel, "Type", "Finish to Start")
            relationship_object_id += 1
        previous_activity_id = current_activity_id

    schedule_options = ET.SubElement(project, qname("ScheduleOptions"))
    add_text(schedule_options, "CalculateFloatBasedOnFinishDate", 1)
    add_text(schedule_options, "ComputeTotalFloatType", "Smallest of Start Float and Finish Float")
    add_text(schedule_options, "CriticalActivityFloatThreshold", 0)
    add_text(schedule_options, "CriticalActivityPathType", "Critical Float")
    add_text(schedule_options, "ExternalProjectPriorityLimit", 5)
    add_text(schedule_options, "IgnoreOtherProjectRelationships", 0)
    add_text(schedule_options, "IncludeExternalResAss", 0)
    add_text(schedule_options, "LevelAllResources", 1)
    add_text(schedule_options, "LevelWithinFloat", 0)
    add_text(schedule_options, "MakeOpenEndedActivitiesCritical", 0)
    add_text(schedule_options, "MaximumMultipleFloatPaths", 10)
    add_text(schedule_options, "MinFloatToPreserve", 1)
    add_text(schedule_options, "MultipleFloatPathsEnabled", 0)
    add_text(schedule_options, "MultipleFloatPathsUseTotalFloat", 1)
    add_text(schedule_options, "OutOfSequenceScheduleType", "Retained Logic")
    add_text(schedule_options, "OverAllocationPercentage", 25)
    add_text(schedule_options, "PreserveScheduledEarlyAndLateDates", 1)
    add_text(schedule_options, "PriorityList", "(0||priority_type(sort_type|ASC)())")
    add_text(schedule_options, "ProjectObjectId", 1)
    add_text(schedule_options, "RelationshipLagCalendar", "24 Hour Calendar")
    add_text(schedule_options, "StartToStartLagCalculationType", 1)
    add_text(schedule_options, "UseExpectedFinishDates", 0)

    return ET.ElementTree(root)


def write_revit_build_code_map(activities: pd.DataFrame) -> None:
    pushback_rows = activities[activities["element_key"].notna() & activities["source_model"].notna()]
    pushback_map = (
        pushback_rows.loc[
            :,
            [
                "element_id",
                "export_task_name",
                "task_id",
                "task_name",
                "level",
                "element_key",
                "element_start",
                "element_end",
            ],
        ]
        .assign(task_name=lambda df: df["export_task_name"])
        .rename(columns={"export_task_name": "build_code"})
        .sort_values(["task_id", "level", "element_start", "element_id"])
        .reset_index(drop=True)
    )
    pushback_map.to_csv(REVIT_BUILD_CODE_MAP_PATH, index=False)


if __name__ == "__main__":
    OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)
    activities = build_activity_table()
    tree = build_xml()
    ET.indent(tree, space="  ")
    tree.write(OUTPUT_XML_PATH, encoding="utf-8", xml_declaration=True)
    write_revit_build_code_map(activities)
    print(f"Wrote {OUTPUT_XML_PATH.name}")
    print(f"Wrote {REVIT_BUILD_CODE_MAP_PATH.name}")
