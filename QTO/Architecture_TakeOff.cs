using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace QTO
{
    [Transaction(TransactionMode.Manual)]
    public class Architecture_TakeOff : IExternalCommand
    {
        private static readonly string[] SnapshotKeywords =
        {
            "size",
            "diameter",
            "radius",
            "width",
            "height",
            "length",
            "area",
            "volume",
            "material",
            "weight",
            "mass",
            "thickness",
            "depth",
            "mark",
            "level",
            "offset",
            "comment",
            "assembly",
            "type",
            "fire",
            "finish"
        };

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            try
            {
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;

                IList<Element> architecturalElements = GetAllArchitecturalElements(doc);
                if (architecturalElements.Count == 0)
                {
                    TaskDialog.Show("Revit", "No architectural elements found in this model.");
                    return Result.Succeeded;
                }

                string csvPath = ExportPathHelper.GetScheduleFilePath(doc, "Architecture_TakeOff");
                ExportElementsToCsv(doc, architecturalElements, csvPath);

                TaskDialog.Show(
                    "Revit Export",
                    $"Exported {architecturalElements.Count} architectural elements to:\n{csvPath}"
                );

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("Error", ex.ToString());
                return Result.Failed;
            }
        }

        private IList<Element> GetAllArchitecturalElements(Document doc)
        {
            List<BuiltInCategory> categories = new List<BuiltInCategory>
            {
                BuiltInCategory.OST_Walls,
                BuiltInCategory.OST_Doors,
                BuiltInCategory.OST_Windows,
                BuiltInCategory.OST_Floors,
                BuiltInCategory.OST_Roofs,
                BuiltInCategory.OST_Ceilings,
                BuiltInCategory.OST_CurtainWallPanels,
                BuiltInCategory.OST_CurtainWallMullions,
                BuiltInCategory.OST_Stairs,
                BuiltInCategory.OST_StairsRuns,
                BuiltInCategory.OST_StairsLandings,
                BuiltInCategory.OST_Railings,
                BuiltInCategory.OST_GenericModel,
                BuiltInCategory.OST_Casework,
                BuiltInCategory.OST_Furniture,
                BuiltInCategory.OST_FurnitureSystems,
                BuiltInCategory.OST_PlumbingFixtures
            };

            ElementMulticategoryFilter filter = new ElementMulticategoryFilter(categories);

            return new FilteredElementCollector(doc)
                .WherePasses(filter)
                .WhereElementIsNotElementType()
                .ToElements()
                .GroupBy(e => e.Id.Value)
                .Select(g => g.First())
                .ToList();
        }

        private void ExportElementsToCsv(Document doc, IList<Element> elementsToExport, string filePath)
        {
            StringBuilder csv = new StringBuilder();
            csv.AppendLine(
                "ElementId,Category,Family,Type,Level,Mark,Assembly Code,Assembly Description,Length,Width,Depth,Height,Area,Volume,Weight,Unit Weight,Material,Type Comments,Base Level,Top Level,Base Offset,Top Offset,Location Type,Position X (ft),Position Y (ft),Position Z (ft),Start X (ft),Start Y (ft),Start Z (ft),End X (ft),End Y (ft),End Z (ft),Rotation (deg),Bounding Box Min X (ft),Bounding Box Min Y (ft),Bounding Box Min Z (ft),Bounding Box Max X (ft),Bounding Box Max Y (ft),Bounding Box Max Z (ft),Bounding Box Center X (ft),Bounding Box Center Y (ft),Bounding Box Center Z (ft),Comments,Parameter Snapshot"
            );

            foreach (Element elem in elementsToExport)
            {
                SpatialElementData spatialData = SpatialElementData.FromElement(elem);
                string elementId = elem.Id.Value.ToString();
                string category = elem.Category?.Name ?? "";
                string family = GetFamilyName(elem);
                string typeName = GetTypeName(doc, elem);
                string level = GetLevelName(doc, elem);
                string mark = GetFirstAvailableParameterValue(doc, elem, "Mark");
                string assemblyCode = GetFirstAvailableParameterValue(doc, elem, "Assembly Code");
                string assemblyDescription = GetFirstAvailableParameterValue(doc, elem, "Assembly Description");
                string length = GetFirstAvailableParameterValue(doc, elem, "Length", "Cut Length", "Span");
                string width = GetFirstAvailableParameterValue(doc, elem, "Width", "Actual Width");
                string depth = GetFirstAvailableParameterValue(
                    doc,
                    elem,
                    "Depth",
                    "Thickness",
                    "Structural Depth"
                );
                string height = GetFirstAvailableParameterValue(doc, elem, "Height", "Thickness");
                string area = GetFirstAvailableParameterValue(
                    doc,
                    elem,
                    "Area",
                    "Host Area Computed",
                    "Computed Area"
                );
                string volume = GetFirstAvailableParameterValue(
                    doc,
                    elem,
                    "Volume",
                    "Host Volume Computed"
                );
                string weight = GetFirstAvailableParameterValue(
                    doc,
                    elem,
                    "Weight",
                    "Calculated Weight",
                    "Mass"
                );
                string unitWeight = GetFirstAvailableParameterValue(
                    doc,
                    elem,
                    "Material: Unit weight",
                    "Unit Weight",
                    "Weight per Unit Length",
                    "Mass per Unit Length"
                );
                string material = GetMaterialSummary(doc, elem);
                string typeComments = GetTypeParameterValue(doc, elem, "Type Comments");
                string baseLevel = GetFirstAvailableParameterValue(doc, elem, "Base Level");
                string topLevel = GetFirstAvailableParameterValue(doc, elem, "Top Level");
                string baseOffset = GetFirstAvailableParameterValue(doc, elem, "Base Offset");
                string topOffset = GetFirstAvailableParameterValue(doc, elem, "Top Offset");
                string comments = GetFirstAvailableParameterValue(doc, elem, "Comments");
                string parameterSnapshot = BuildParameterSnapshot(doc, elem);

                csv.AppendLine(string.Join(",",
                    EscapeCsv(elementId),
                    EscapeCsv(category),
                    EscapeCsv(family),
                    EscapeCsv(typeName),
                    EscapeCsv(level),
                    EscapeCsv(mark),
                    EscapeCsv(assemblyCode),
                    EscapeCsv(assemblyDescription),
                    EscapeCsv(length),
                    EscapeCsv(width),
                    EscapeCsv(depth),
                    EscapeCsv(height),
                    EscapeCsv(area),
                    EscapeCsv(volume),
                    EscapeCsv(weight),
                    EscapeCsv(unitWeight),
                    EscapeCsv(material),
                    EscapeCsv(typeComments),
                    EscapeCsv(baseLevel),
                    EscapeCsv(topLevel),
                    EscapeCsv(baseOffset),
                    EscapeCsv(topOffset),
                    EscapeCsv(spatialData.LocationType),
                    EscapeCsv(spatialData.PositionXFeet),
                    EscapeCsv(spatialData.PositionYFeet),
                    EscapeCsv(spatialData.PositionZFeet),
                    EscapeCsv(spatialData.StartXFeet),
                    EscapeCsv(spatialData.StartYFeet),
                    EscapeCsv(spatialData.StartZFeet),
                    EscapeCsv(spatialData.EndXFeet),
                    EscapeCsv(spatialData.EndYFeet),
                    EscapeCsv(spatialData.EndZFeet),
                    EscapeCsv(spatialData.RotationDegrees),
                    EscapeCsv(spatialData.BoundingBoxMinXFeet),
                    EscapeCsv(spatialData.BoundingBoxMinYFeet),
                    EscapeCsv(spatialData.BoundingBoxMinZFeet),
                    EscapeCsv(spatialData.BoundingBoxMaxXFeet),
                    EscapeCsv(spatialData.BoundingBoxMaxYFeet),
                    EscapeCsv(spatialData.BoundingBoxMaxZFeet),
                    EscapeCsv(spatialData.BoundingBoxCenterXFeet),
                    EscapeCsv(spatialData.BoundingBoxCenterYFeet),
                    EscapeCsv(spatialData.BoundingBoxCenterZFeet),
                    EscapeCsv(comments),
                    EscapeCsv(parameterSnapshot)
                ));
            }

            File.WriteAllText(filePath, csv.ToString(), Encoding.UTF8);
        }

        private string BuildParameterSnapshot(Document doc, Element elem)
        {
            Dictionary<string, string> values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (Parameter parameter in elem.Parameters.Cast<Parameter>())
            {
                AddSnapshotValue(values, parameter, doc);
            }

            ElementId typeId = elem.GetTypeId();
            if (typeId != ElementId.InvalidElementId)
            {
                Element? typeElem = doc.GetElement(typeId);
                if (typeElem != null)
                {
                    foreach (Parameter parameter in typeElem.Parameters.Cast<Parameter>())
                    {
                        AddSnapshotValue(values, parameter, doc);
                    }
                }
            }

            return string.Join(
                " | ",
                values
                    .OrderBy(kvp => kvp.Key, StringComparer.OrdinalIgnoreCase)
                    .Select(kvp => $"{kvp.Key}={kvp.Value}")
            );
        }

        private void AddSnapshotValue(Dictionary<string, string> values, Parameter parameter, Document doc)
        {
            string name = parameter.Definition?.Name ?? "";
            if (string.IsNullOrWhiteSpace(name))
                return;

            string lowered = name.ToLowerInvariant();
            if (!SnapshotKeywords.Any(keyword => lowered.Contains(keyword)))
                return;

            string value = GetParameterValue(parameter, doc);
            if (string.IsNullOrWhiteSpace(value))
                return;

            values.TryAdd(name, value);
        }

        private string GetFamilyName(Element elem)
        {
            if (elem is FamilyInstance fi && fi.Symbol?.Family != null)
            {
                return fi.Symbol.Family.Name;
            }

            return "";
        }

        private string GetTypeName(Document doc, Element elem)
        {
            ElementId typeId = elem.GetTypeId();
            if (typeId != ElementId.InvalidElementId)
            {
                Element typeElem = doc.GetElement(typeId);
                return typeElem?.Name ?? "";
            }

            return "";
        }

        private string GetLevelName(Document doc, Element elem)
        {
            Parameter levelParam = elem.LookupParameter("Level");
            if (levelParam != null)
            {
                return GetParameterValue(levelParam, doc);
            }

            if (elem.LevelId != ElementId.InvalidElementId)
            {
                Element levelElem = doc.GetElement(elem.LevelId);
                return levelElem?.Name ?? "";
            }

            return "";
        }

        private string GetFirstAvailableParameterValue(Document doc, Element elem, params string[] parameterNames)
        {
            foreach (string parameterName in parameterNames)
            {
                string value = GetParameterValue(elem.LookupParameter(parameterName), doc);
                if (!string.IsNullOrWhiteSpace(value))
                    return value;

                ElementId typeId = elem.GetTypeId();
                if (typeId == ElementId.InvalidElementId)
                    continue;

                Element? typeElem = doc.GetElement(typeId);
                if (typeElem == null)
                    continue;

                value = GetParameterValue(typeElem.LookupParameter(parameterName), doc);
                if (!string.IsNullOrWhiteSpace(value))
                    return value;
            }

            return "";
        }

        private string GetTypeParameterValue(Document doc, Element elem, string parameterName)
        {
            ElementId typeId = elem.GetTypeId();
            if (typeId == ElementId.InvalidElementId)
                return "";

            Element typeElem = doc.GetElement(typeId);
            if (typeElem == null)
                return "";

            return GetParameterValue(typeElem.LookupParameter(parameterName), doc);
        }

        private string GetMaterialSummary(Document doc, Element elem)
        {
            ICollection<ElementId> materialIds = elem.GetMaterialIds(false);
            if (materialIds == null || materialIds.Count == 0)
                return "";

            return string.Join("; ",
                materialIds
                    .Select(doc.GetElement)
                    .OfType<Material>()
                    .Select(material => material.Name)
                    .Distinct()
            );
        }

        private string GetParameterValue(Parameter para, Document document)
        {
            if (para == null)
                return "";

            try
            {
                switch (para.StorageType)
                {
                    case StorageType.Double:
                        return para.AsValueString() ?? "";

                    case StorageType.ElementId:
                        ElementId id = para.AsElementId();
                        if (id == ElementId.InvalidElementId)
                            return "";

                        if (id.Value >= 0)
                        {
                            Element referencedElem = document.GetElement(id);
                            return referencedElem?.Name ?? id.Value.ToString();
                        }

                        return id.Value.ToString();

                    case StorageType.Integer:
                        if (para.Definition != null &&
                            para.Definition.GetDataType() == SpecTypeId.Boolean.YesNo)
                        {
                            return para.AsInteger() == 0 ? "False" : "True";
                        }

                        return para.AsInteger().ToString();

                    case StorageType.String:
                        return para.AsString() ?? "";

                    case StorageType.None:
                    default:
                        return "";
                }
            }
            catch
            {
                return "";
            }
        }

        private string EscapeCsv(string value)
        {
            if (value == null)
                return "";

            value = value.Replace("\"", "\"\"");

            if (value.Contains(",") || value.Contains("\"") || value.Contains("\n") || value.Contains("\r"))
            {
                return $"\"{value}\"";
            }

            return value;
        }
    }
}




