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
