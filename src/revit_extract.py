import json
import os
import sys

import clr

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import BuiltInCategory, FilteredElementCollector, ProjectInfo


try:
    uiapp = __revit__
except NameError:
    raise Exception(
        "This script must run inside Revit through pyRevit or RevitPythonShell. "
        "A normal standalone Python interpreter cannot load and drive the Revit API by itself."
    )


def get_document_from_args(app):
    if len(sys.argv) > 1:
        model_path = sys.argv[1]
        return app.OpenDocumentFile(model_path)

    active_uidoc = uiapp.ActiveUIDocument
    if active_uidoc is None:
        raise Exception("Open a model in Revit or pass a model path as the first script argument.")
    return active_uidoc.Document


def get_project_info(doc):
    info = doc.ProjectInformation
    if info is None or not isinstance(info, ProjectInfo):
        return {}

    return {
        "name": info.Name,
        "number": info.Number,
        "address": info.Address,
        "client_name": info.ClientName,
        "status": info.Status,
    }


def get_counts(doc):
    category_map = [
        ("walls", BuiltInCategory.OST_Walls),
        ("doors", BuiltInCategory.OST_Doors),
        ("windows", BuiltInCategory.OST_Windows),
        ("rooms", BuiltInCategory.OST_Rooms),
        ("floors", BuiltInCategory.OST_Floors),
        ("structural_columns", BuiltInCategory.OST_StructuralColumns),
    ]

    counts = {}
    for label, bic in category_map:
        elements = FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType()
        counts[label] = elements.GetElementCount()
    return counts


def get_level_names(doc):
    levels = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType()
    names = []
    for level in levels:
        names.append(level.Name)
    return sorted(names)


def main():
    app = uiapp.Application
    doc = get_document_from_args(app)
    close_when_done = doc != getattr(uiapp.ActiveUIDocument, "Document", None)

    payload = {
        "title": doc.Title,
        "path_name": doc.PathName,
        "project_info": get_project_info(doc),
        "element_counts": get_counts(doc),
        "levels": get_level_names(doc),
    }

    output_path = os.environ.get("REVIT_EXTRACT_OUTPUT", os.path.join(os.getcwd(), "revit_extract.json"))
    with open(output_path, "w") as handle:
        json.dump(payload, handle, indent=2)

    if close_when_done:
        doc.Close(False)

    sys.stdout.write("Wrote {0}\n".format(output_path))


if __name__ == "__main__":
    main()
