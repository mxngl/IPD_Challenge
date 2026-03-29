using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace QTO
{
    [Transaction(TransactionMode.Manual)]
    public class Structural_TakeOff : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            try
            {
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;

                IList<Element> structuralElements = GetAllStructuralElements(doc);

                if (structuralElements.Count == 0)
                {
                    TaskDialog.Show("Revit", "No structural elements found in this model.");
                    return Result.Succeeded;
                }

                string csvPath = GetScheduleFilePath();

                ExportElementsToCsv(doc, structuralElements, csvPath);

                TaskDialog.Show(
                    "Revit Export",
                    $"Exported {structuralElements.Count} structural elements to:\n{csvPath}"
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

        private IList<Element> GetAllStructuralElements(Document doc)
        {
            List<BuiltInCategory> categories = new List<BuiltInCategory>
            {
                BuiltInCategory.OST_StructuralColumns,
                BuiltInCategory.OST_StructuralFraming,
                BuiltInCategory.OST_StructuralFoundation,
                BuiltInCategory.OST_StructuralStiffener,
                BuiltInCategory.OST_StructuralTruss,
                BuiltInCategory.OST_StructConnections,
                BuiltInCategory.OST_StructConnectionPlates,
                BuiltInCategory.OST_StructConnectionBolts,
                BuiltInCategory.OST_StructConnectionAnchors,
                BuiltInCategory.OST_Rebar,
                BuiltInCategory.OST_AreaRein,
                BuiltInCategory.OST_PathRein,
                BuiltInCategory.OST_FabricReinforcement
            };

            ElementMulticategoryFilter filter = new ElementMulticategoryFilter(
                categories.Cast<BuiltInCategory>().ToList()
            );

            return new FilteredElementCollector(doc)
                .WherePasses(filter)
                .WhereElementIsNotElementType()
                .ToElements()
                .GroupBy(e => e.Id.Value)
                .Select(g => g.First())
                .ToList();
        }

        private string GetScheduleFilePath()
        {
            string assemblyDirectory = Path.GetDirectoryName(
                System.Reflection.Assembly.GetExecutingAssembly().Location
            ) ?? "";

            DirectoryInfo? directory = new DirectoryInfo(assemblyDirectory);
            while (directory != null &&
                   !Directory.Exists(Path.Combine(directory.FullName, "revit_schedules")))
            {
                directory = directory.Parent;
            }

            string schedulesDirectory = directory != null
                ? Path.Combine(directory.FullName, "revit_schedules")
                : Path.Combine(assemblyDirectory, "revit_schedules");

            Directory.CreateDirectory(schedulesDirectory);

            return Path.Combine(schedulesDirectory, "Structural_Schedule.csv");
        }

        private void ExportElementsToCsv(Document doc, IList<Element> elementsToExport, string filePath)
        {
            StringBuilder csv = new StringBuilder();

            // Header row
            csv.AppendLine(
                "ElementId,Category,Family,Type,Level,Mark,Assembly Code,Assembly Description,Length,Width,Depth,Height,Area,Volume,Material,Type Comments,Base Level,Top Level,Base Offset,Top Offset,Comments"
            );

            foreach (Element elem in elementsToExport)
            {
                string elementId = elem.Id.Value.ToString();
                string category = elem.Category?.Name ?? "";
                string family = GetFamilyName(elem);
                string typeName = GetTypeName(doc, elem);
                string level = GetLevelName(doc, elem);
                string mark = GetParameterValue(elem.LookupParameter("Mark"), doc);
                string assemblyCode = GetAssemblyCode(elem, doc);
                string assemblyDescription = GetParameterValue(elem.LookupParameter("Assembly Description"), doc);
                string length = GetParameterValue(elem.LookupParameter("Length"), doc);
                string width = GetParameterValue(elem.LookupParameter("Width"), doc);
                string depth = GetParameterValue(elem.LookupParameter("Depth"), doc);
                string height = GetParameterValue(elem.LookupParameter("Height"), doc);
                string area = GetParameterValue(elem.LookupParameter("Area"), doc);
                string volume = GetParameterValue(elem.LookupParameter("Volume"), doc);
                string material = GetMaterialSummary(doc, elem);
                string typeComments = GetTypeParameterValue(doc, elem, "Type Comments");
                string baseLevel = GetParameterValue(elem.LookupParameter("Base Level"), doc);
                string topLevel = GetParameterValue(elem.LookupParameter("Top Level"), doc);
                string baseOffset = GetParameterValue(elem.LookupParameter("Base Offset"), doc);
                string topOffset = GetParameterValue(elem.LookupParameter("Top Offset"), doc);
                string comments = GetParameterValue(elem.LookupParameter("Comments"), doc);

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
                    EscapeCsv(material),
                    EscapeCsv(typeComments),
                    EscapeCsv(baseLevel),
                    EscapeCsv(topLevel),
                    EscapeCsv(baseOffset),
                    EscapeCsv(topOffset),
                    EscapeCsv(comments)
                ));
            }

            File.WriteAllText(filePath, csv.ToString(), Encoding.UTF8);
        }

        private string GetFamilyName(Element elem)
        {
            if (elem is FamilyInstance fi && fi.Symbol != null && fi.Symbol.Family != null)
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

        private string GetAssemblyCode(Element elem, Document doc)
        {
            Parameter parameter = elem.LookupParameter("Assembly Code");
            string value = GetParameterValue(parameter, doc);

            if (!string.IsNullOrWhiteSpace(value))
                return value;

            ElementId typeId = elem.GetTypeId();
            if (typeId == ElementId.InvalidElementId)
                return "";

            Element typeElem = doc.GetElement(typeId);
            if (typeElem == null)
                return "";

            Parameter typeParameter = typeElem.LookupParameter("Assembly Code");
            value = GetParameterValue(typeParameter, doc);

            return string.IsNullOrWhiteSpace(value) ? "" : value;
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
                        else
                        {
                            return id.Value.ToString();
                        }

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

