# Autodesk Cloud -> Revit Extraction

This scaffold splits the workflow into two parts because these are different runtimes:

1. `src/aps_rvt_download.py`
   Connects to Autodesk Platform Services (APS, formerly Forge) and downloads a specific ACC/BIM 360 file version as a local `.rvt` file.
2. `src/revit_extract.py`
   Runs inside Revit through pyRevit or RevitPythonShell and uses the Revit API to extract data from the `.rvt` model into JSON.

## Why two scripts are necessary

The Revit API is not a normal standalone Python library. It must run in a Revit-hosted process such as:

- Revit + pyRevit
- Revit + RevitPythonShell
- Revit Design Automation
- a compiled Revit add-in

So the practical flow is:

1. Download the file from Autodesk cloud.
2. Open that file in a Revit-hosted environment.
3. Run the extraction script there.

## Setup

```powershell
python -m pip install -r requirements.txt
```

Set APS credentials:

```powershell
$env:APS_CLIENT_ID="your-client-id"
$env:APS_CLIENT_SECRET="your-client-secret"
```

## STV workflow

This repo also includes a Python-based STV engine that can:

- read structural and MEP takeoff CSVs exported from Revit
- map those schedule rows into STV construction items
- calculate STV impacts using the workbook in `STV_Template/`
- generate JSON outputs and PNG charts under `outputs/`

The main CLI entrypoint is `src/STV_Engine/cli.py`.

### Run the structural STV flow

From the repo root:

```powershell
python -m src.STV_Engine.cli --team Island --structural-schedule .\revit_schedules\Structural_Schedule.csv --output-dir .\outputs\stv_structural
```

This command:

- reads `revit_schedules/Structural_Schedule.csv`
- converts recognized rows into STV construction items
- uses `STV_Template/STV_ConceptA_Bambo.xlsx` as the default reference workbook
- writes the structural report and plots into `outputs/stv_structural`

Generated files include:

- `outputs/stv_structural/stv_results.json`
- `outputs/stv_structural/structural_schedule_items.json`
- `outputs/stv_structural/target_vs_project.png`
- `outputs/stv_structural/life_cycle_breakdown.png`
- `outputs/stv_structural/target_vs_project_radial.png`

If you want to explicitly point at a workbook, add:

```powershell
--template .\STV_Template\STV_ConceptA_Bambo.xlsx
```

### Run the MEP STV flow

```powershell
python -m src.STV_Engine.cli --team Island --mep-schedule .\revit_schedules\MEP_TakeOff.csv --output-dir .\outputs\stv_mep
```

This writes:

- `outputs/stv_mep/stv_results.json`
- `outputs/stv_mep/mep_schedule_items.json`
- `outputs/stv_mep/target_vs_project.png`
- `outputs/stv_mep/life_cycle_breakdown.png`
- `outputs/stv_mep/target_vs_project_radial.png`

### MEP mapping notes

The MEP importer in `src/STV_Engine/revit_mep.py` uses explicit `Family : Category` mappings for the current Revit takeoff export.

Current mappings include:

- `Supply Diffuser : Air Terminals` -> `Air Handling Unit (m^3/s)`
- `Exhaust Grill : Air Terminals` -> `Air Handling Unit (m^3/s)`
- `PRICE-40FF- Filter Frame Stamped Residential Grille-RETURN Hosted : Air Terminals` -> `Air Filters (m^3/s)`
- `Return Diffuser : Air Terminals` -> `Air Handling Unit (m^3/s)`
- `Outdoor AHU - Horizontal : Mechanical Equipment` -> `Air Handling Unit (m^3/s)`
- `Rectangular Elbow - Mitered : Duct Fittings` -> `Stainless Steel Duct (kg)`
- `Rectangular Tee : Duct Fittings` -> `Stainless Steel Duct (kg)`
- `Rectangular Cross : Duct Fittings` -> `Stainless Steel Duct (kg)`
- `Rectangular Transition - Angle : Duct Fittings` -> `Stainless Steel Duct (kg)`
- `Rectangular to Round Transition - Angle : Duct Fittings` -> `Stainless Steel Duct (kg)`

Currently skipped on purpose:

- `34274 : Electrical Fixtures`
- `Utility Switchboard : Electrical Equipment`
- rows with missing `Family`

### Duct weight approximation

For stainless rectangular duct fittings, the current workflow uses a practical early-stage mass estimate when the Revit export does not include a usable `Weight` or `Unit Weight` value.

Base rule:

```text
Weight ≈ 2 * (W + H) * L * t * rho
```

with:

- all dimensions in meters
- `rho = 8000 kg/m^3`
- sheet thickness selected by size class

Thickness assumptions:

- small ducts: `0.5 mm`
- medium ducts: `0.6 mm`
- large ducts: `0.8 mm`

In the current implementation, the size class is based on the largest duct side:

- `<= 0.3 m` -> `0.5 mm`
- `<= 0.6 m` -> `0.6 mm`
- `> 0.6 m` -> `0.8 mm`

The parser reads duct dimensions from the exported `Width`, `Height`, `Size`, `Length`, and `Parameter Snapshot` fields, then falls back to a simple volume-based estimate for length when needed.

### Fitting multipliers

For rectangular fittings, the current approximation applies a surface-area multiplier on top of the straight-duct estimate:

- mitered elbow -> `1.3`
- tee -> `1.6`
- cross -> `1.8`
- rectangular transition -> `1.2`
- rectangular-to-round transition -> `1.25`

These are practical BIM/LCA assumptions for early-stage estimating, not fabrication-grade quantities.

### Combine individual STVs into a project STV

If you already have separate result folders, you can merge their `stv_results.json` files into one project-level STV:

```powershell
python -m src.STV_Engine.cli --combine-results .\outputs\stv_structural\stv_results.json .\outputs\stv_mep\stv_results.json --output-dir .\outputs\stv_project
```

This writes:

- `outputs/stv_project/stv_results.json`
- `outputs/stv_project/target_vs_project.png`
- `outputs/stv_project/life_cycle_breakdown.png`
- `outputs/stv_project/target_vs_project_radial.png`

The combine step expects the individual STVs to share the same team, target values, and lifetime.

### Use a JSON input file

You can also start from a JSON payload and optionally combine it with schedule-derived items:

```powershell
python -m src.STV_Engine.cli --input .\src\STV_Engine\examples\concept_a_bambo.json --team Island --output-dir .\outputs\stv_example
```

## STV dashboard page

A static dashboard page is available at `index.html`. It loads the existing files in:

- `outputs/stv_project/`
- `outputs/stv_structural/`
- `outputs/stv_mep/`
- `outputs/stv_example/`

and displays the generated plots with a few headline metrics.

### Preview locally

If you open `index.html` directly as a file, some browsers may block `fetch()` calls to local JSON files. The easiest way to preview it is to serve the repo with a simple local web server:

```powershell
python -m http.server 8000
```

Then open:

```text
http://localhost:8000/
```

### Publish with GitHub Pages

To publish the STV plots as a simple hosted page:

1. Push this repository to GitHub.
2. Enable GitHub Pages for the branch that contains `index.html`.
3. Make sure the `outputs/` folder is published alongside `index.html`.
4. Open the resulting GitHub Pages URL.

You can then paste that public URL into a Notion page as an embed or bookmark.

## Step 1: Download a model from Autodesk cloud

You need the ACC/BIM 360 `project_id` and the target `version_id`.

```powershell
python src/aps_rvt_download.py --project-id <project_id> --version-id <version_id> --output .\downloads\model.rvt
```

What this script does:

- authenticates with APS OAuth 2-legged credentials
- fetches version metadata from the Data Management API
- reads the underlying storage URN
- requests a signed download URL from OSS
- downloads the `.rvt` file locally

## Step 2: Extract data with the Revit API

Run `src/revit_extract.py` inside Revit using pyRevit or RevitPythonShell.

If the model is already open in Revit:

```python
exec(open(r"C:\path\to\src\revit_extract.py").read())
```

If you want the script to open a specific model path from inside the Revit-hosted Python environment:

```python
import os
import sys
os.environ["REVIT_EXTRACT_OUTPUT"] = r"C:\temp\revit_extract.json"
sys.argv = ["revit_extract.py", r"C:\path\to\downloads\model.rvt"]
exec(open(r"C:\path\to\src\revit_extract.py").read())
```

The extractor currently writes:

- model title and path
- project information
- counts for walls, doors, windows, rooms, floors, and structural columns
- level names

Output defaults to `revit_extract.json` in the current working directory, or the path set in `REVIT_EXTRACT_OUTPUT`.

## If you want full automation

If your goal is fully unattended processing directly from Autodesk cloud, the next step is usually one of these:

1. Autodesk Revit Design Automation appbundle + activity
2. A desktop machine with Revit installed running this flow on a schedule
3. A compiled Revit add-in instead of a Python script

If you want, I can build the next iteration in one of those directions.
