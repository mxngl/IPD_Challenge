from __future__ import annotations

import argparse
from datetime import datetime, timezone
import json
from pathlib import Path
import re

from .engine import STVEngine
from .revit_architecture import load_architecture_schedule
from .models import STVInputs, STVResults
from .revit_mep import load_mep_schedule
from .reference import DEFAULT_TEMPLATE_PATH, STVReferenceData
from .revit_structural import load_structural_schedule
from .visualization import export_visualizations


ARCHITECTURE_HISTORY_TIMESTAMP_RE = re.compile(
    r"_(\d{4}-\d{2}-\d{2})_(\d{2})-(\d{2})-(\d{2})(am|pm)_",
    re.IGNORECASE,
)


def _history_entry(
    results: STVResults,
    *,
    timestamp: str,
    source_file: str | None = None,
    source_path: str | None = None,
) -> dict[str, object]:
    entry: dict[str, object] = {
        "timestamp": timestamp,
        "team": results.team,
        "lifetime_years": results.lifetime_years,
        "metric_summary": results.metric_summary(),
        "breakdown": results.breakdown.to_dict(),
    }
    if source_file is not None:
        entry["source_file"] = source_file
    if source_path is not None:
        entry["source_path"] = source_path
    return entry


def append_results_history(
    output_dir: Path,
    results: STVResults,
    *,
    previous_results: STVResults | None = None,
    previous_timestamp: str | None = None,
) -> Path:
    history_path = output_dir / "history.json"
    if history_path.exists():
        history = json.loads(history_path.read_text(encoding="utf-8"))
    else:
        history = []
        if previous_results is not None:
            history.append(_history_entry(
                previous_results,
                timestamp=previous_timestamp or datetime.now(timezone.utc).isoformat(),
            ))

    history.append(_history_entry(results, timestamp=datetime.now(timezone.utc).isoformat()))
    history_path.write_text(json.dumps(history, indent=2), encoding="utf-8")
    return history_path


def _parse_schedule_timestamp(schedule_path: Path) -> datetime:
    match = ARCHITECTURE_HISTORY_TIMESTAMP_RE.search(schedule_path.name)
    if match:
        date_part, hour_text, minute_text, second_text, meridiem = match.groups()
        hour = int(hour_text)
        if meridiem.lower() == "am":
            hour = 0 if hour == 12 else hour
        else:
            hour = 12 if hour == 12 else hour + 12

        parsed = datetime.fromisoformat(
            f"{date_part}T{hour:02d}:{minute_text}:{second_text}"
        )
        return parsed.replace(tzinfo=timezone.utc)

    return datetime.fromtimestamp(schedule_path.stat().st_mtime, tz=timezone.utc)


def _run_stv(
    payload: dict[str, object],
    *,
    team: str,
    template_path: str,
) -> STVResults:
    payload["team"] = team
    inputs = STVInputs.from_dict(payload)
    reference_data = STVReferenceData.from_workbook(template_path)
    engine = STVEngine(reference_data)
    return engine.calculate(inputs)


def _run_architecture_history(
    schedule_dir: Path,
    *,
    team: str,
    output_dir: Path,
    template_path: str,
) -> dict[str, object]:
    schedule_paths = sorted(
        schedule_dir.glob("*.csv"),
        key=_parse_schedule_timestamp,
    )
    if not schedule_paths:
        raise ValueError(f"No architecture schedule CSVs found in {schedule_dir}.")

    history: list[dict[str, object]] = []
    latest_report = None
    latest_results = None
    latest_schedule_path = None

    for schedule_path in schedule_paths:
        report = load_architecture_schedule(schedule_path)
        payload = {
            "construction_items": [
                {
                    "assembly": item.assembly,
                    "material_type": item.material_type,
                    "amount": item.amount,
                }
                for item in report.construction_items
            ]
        }
        results = _run_stv(payload, team=team, template_path=template_path)
        timestamp = _parse_schedule_timestamp(schedule_path).isoformat()
        history.append(
            _history_entry(
                results,
                timestamp=timestamp,
                source_file=schedule_path.name,
                source_path=str(schedule_path),
            )
        )
        latest_report = report
        latest_results = results
        latest_schedule_path = schedule_path

    assert latest_report is not None
    assert latest_results is not None
    assert latest_schedule_path is not None

    results_path = output_dir / "stv_results.json"
    results_path.write_text(
        json.dumps(latest_results.to_dict(), indent=2),
        encoding="utf-8",
    )

    architecture_report_path = output_dir / "architecture_schedule_items.json"
    architecture_report_path.write_text(
        json.dumps(latest_report.to_dict(), indent=2),
        encoding="utf-8",
    )

    history_path = output_dir / "history.json"
    history_path.write_text(json.dumps(history, indent=2), encoding="utf-8")

    image_paths = export_visualizations(latest_results, output_dir)
    return {
        "results_json": str(results_path),
        "charts": {key: str(path) for key, path in image_paths.items()},
        "history_json": str(history_path),
        "architecture_schedule_items": str(architecture_report_path),
        "history_source_dir": str(schedule_dir),
        "history_source_files": [str(path) for path in schedule_paths],
        "latest_schedule": str(latest_schedule_path),
    }


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Run Sustainable Target Value calculations.")
    parser.add_argument("--input", help="Path to JSON inputs.")
    parser.add_argument(
        "--template",
        default=str(DEFAULT_TEMPLATE_PATH),
        help="Path to the STV Excel workbook used as the reference dataset.",
    )
    parser.add_argument("--team", help="Team name, for example Island.")
    parser.add_argument(
        "--structural-schedule",
        help="Path to a Revit structural schedule CSV to convert into embodied STV inputs.",
    )
    parser.add_argument(
        "--mep-schedule",
        help="Path to a Revit MEP takeoff CSV to convert into embodied STV inputs.",
    )
    parser.add_argument(
        "--architecture-schedule",
        help="Path to a Revit architecture takeoff CSV to convert into embodied STV inputs.",
    )
    parser.add_argument(
        "--architecture-history-dir",
        help="Directory of Revit architecture takeoff CSVs to batch into a time-history STV run.",
    )
    parser.add_argument(
        "--output-dir",
        default="outputs/stv",
        help="Directory for result JSON and generated charts.",
    )
    parser.add_argument(
        "--combine-results",
        nargs="+",
        help="Paths to one or more existing stv_results.json files to merge into a project STV.",
    )
    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    results_path = output_dir / "stv_results.json"
    previous_results = None
    previous_timestamp = None
    if results_path.exists():
        previous_results = STVResults.from_dict(json.loads(results_path.read_text(encoding="utf-8")))
        previous_timestamp = datetime.fromtimestamp(
            results_path.stat().st_mtime,
            tz=timezone.utc,
        ).isoformat()

    if args.combine_results:
        result_paths = [Path(path) for path in args.combine_results]
        loaded_results = [
            STVResults.from_dict(json.loads(path.read_text(encoding="utf-8")))
            for path in result_paths
        ]
        combined_results = STVResults.combine(loaded_results, team=args.team)

        results_path.write_text(
            json.dumps(combined_results.to_dict(), indent=2),
            encoding="utf-8",
        )
        image_paths = export_visualizations(combined_results, output_dir)
        history_path = append_results_history(
            output_dir,
            combined_results,
            previous_results=previous_results,
            previous_timestamp=previous_timestamp,
        )

        response: dict[str, object] = {
            "results_json": str(results_path),
            "charts": {key: str(path) for key, path in image_paths.items()},
            "combined_from": [str(path) for path in result_paths],
            "history_json": str(history_path),
        }
        print(json.dumps(response, indent=2))
        return

    if args.architecture_history_dir:
        team = args.team
        if not team:
            parser.error("Provide a team with --team when using --architecture-history-dir.")
        response = _run_architecture_history(
            Path(args.architecture_history_dir),
            team=team,
            output_dir=output_dir,
            template_path=args.template,
        )
        print(json.dumps(response, indent=2))
        return

    payload: dict[str, object] = {}
    if args.input:
        input_path = Path(args.input)
        payload = json.loads(input_path.read_text(encoding="utf-8"))

    if args.structural_schedule:
        report = load_structural_schedule(args.structural_schedule)
        existing_items = list(payload.get("construction_items", []))
        schedule_items = [
            {
                "assembly": item.assembly,
                "material_type": item.material_type,
                "amount": item.amount,
            }
            for item in report.construction_items
        ]
        payload["construction_items"] = existing_items + schedule_items
        structural_report_path = output_dir / "structural_schedule_items.json"
        structural_report_path.write_text(
            json.dumps(report.to_dict(), indent=2),
            encoding="utf-8",
        )

    if args.mep_schedule:
        report = load_mep_schedule(args.mep_schedule)
        existing_items = list(payload.get("construction_items", []))
        mep_items = [
            {
                "assembly": item.assembly,
                "material_type": item.material_type,
                "amount": item.amount,
            }
            for item in report.construction_items
        ]
        payload["construction_items"] = existing_items + mep_items
        mep_report_path = output_dir / "mep_schedule_items.json"
        mep_report_path.write_text(
            json.dumps(report.to_dict(), indent=2),
            encoding="utf-8",
        )

    if args.architecture_schedule:
        report = load_architecture_schedule(args.architecture_schedule)
        existing_items = list(payload.get("construction_items", []))
        architecture_items = [
            {
                "assembly": item.assembly,
                "material_type": item.material_type,
                "amount": item.amount,
            }
            for item in report.construction_items
        ]
        payload["construction_items"] = existing_items + architecture_items
        architecture_report_path = output_dir / "architecture_schedule_items.json"
        architecture_report_path.write_text(
            json.dumps(report.to_dict(), indent=2),
            encoding="utf-8",
        )

    team = args.team or payload.get("team")
    if not team:
        parser.error("Provide a team with --team or in the input JSON.")
    results = _run_stv(payload, team=team, template_path=args.template)

    results_path.write_text(
        json.dumps(results.to_dict(), indent=2),
        encoding="utf-8",
    )
    image_paths = export_visualizations(results, output_dir)
    history_path = append_results_history(
        output_dir,
        results,
        previous_results=previous_results,
        previous_timestamp=previous_timestamp,
    )

    response: dict[str, object] = {
        "results_json": str(results_path),
        "charts": {key: str(path) for key, path in image_paths.items()},
        "history_json": str(history_path),
    }
    if args.structural_schedule:
        response["structural_schedule_items"] = str(output_dir / "structural_schedule_items.json")
    if args.mep_schedule:
        response["mep_schedule_items"] = str(output_dir / "mep_schedule_items.json")
    if args.architecture_schedule:
        response["architecture_schedule_items"] = str(output_dir / "architecture_schedule_items.json")
    print(json.dumps(response, indent=2))


if __name__ == "__main__":
    main()
