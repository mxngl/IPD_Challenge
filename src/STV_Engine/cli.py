from __future__ import annotations

import argparse
from datetime import datetime, timezone
import json
from pathlib import Path

from .engine import STVEngine
from .revit_architecture import load_architecture_schedule
from .models import STVInputs, STVResults
from .revit_mep import load_mep_schedule
from .reference import DEFAULT_TEMPLATE_PATH, STVReferenceData
from .revit_structural import load_structural_schedule
from .visualization import export_visualizations


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
            history.append(
                {
                    "timestamp": previous_timestamp or datetime.now(timezone.utc).isoformat(),
                    "team": previous_results.team,
                    "lifetime_years": previous_results.lifetime_years,
                    "metric_summary": previous_results.metric_summary(),
                    "breakdown": previous_results.breakdown.to_dict(),
                }
            )

    summary = results.metric_summary()
    history.append(
        {
            "timestamp": datetime.now(timezone.utc).isoformat(),
            "team": results.team,
            "lifetime_years": results.lifetime_years,
            "metric_summary": summary,
            "breakdown": results.breakdown.to_dict(),
        }
    )
    history_path.write_text(json.dumps(history, indent=2), encoding="utf-8")
    return history_path


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
    payload["team"] = team

    inputs = STVInputs.from_dict(payload)
    reference_data = STVReferenceData.from_workbook(args.template)
    engine = STVEngine(reference_data)
    results = engine.calculate(inputs)

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
