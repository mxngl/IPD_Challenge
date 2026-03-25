using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace HelloWorld
{
    [Transaction(TransactionMode.Manual)]
    public class HelloCommand : IExternalCommand
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

                IList<Element> columns = GetAllColumns(doc);

                if (columns.Count == 0)
                {
                    TaskDialog.Show("Revit", "No columns found in this model.");
                    return Result.Succeeded;
                }

                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string csvPath = Path.Combine(desktopPath, "revit_columns_export.csv");

                ExportColumnsToCsv(doc, columns, csvPath);

                TaskDialog.Show(
                    "Revit Export",
                    $"Exported {columns.Count} columns to:\n{csvPath}"
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

        private IList<Element> GetAllColumns(Document doc)
        {
            List<Element> allColumns = new List<Element>();

            // Structural columns
            FilteredElementCollector structuralCollector = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_StructuralColumns)
                .WhereElementIsNotElementType();

            allColumns.AddRange(structuralCollector.ToElements());

            // Architectural columns
            FilteredElementCollector archCollector = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Columns)
                .WhereElementIsNotElementType();

            allColumns.AddRange(archCollector.ToElements());

            // Remove duplicates if any
            return allColumns
                .GroupBy(e => e.Id.Value)
                .Select(g => g.First())
                .ToList();
        }

        private void ExportColumnsToCsv(Document doc, IList<Element> columns, string filePath)
        {
            StringBuilder csv = new StringBuilder();

            // Header row
            csv.AppendLine(
                "ElementId,Category,Family,Type,Level,Mark,Base Level,Top Level,Base Offset,Top Offset,Comments"
            );

            foreach (Element elem in columns)
            {
                string elementId = elem.Id.Value.ToString();
                string category = elem.Category?.Name ?? "";
                string family = GetFamilyName(elem);
                string typeName = GetTypeName(doc, elem);
                string level = GetLevelName(doc, elem);
                string mark = GetParameterValue(elem.LookupParameter("Mark"), doc);
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
