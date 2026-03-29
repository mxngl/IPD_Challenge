using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;

namespace QTO
{
    [Transaction(TransactionMode.Manual)]
    public class MEP_TakeOff : IExternalCommand
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
            "insulation",
            "lining",
            "flow",
            "airflow",
            "pressure",
            "capacity",
            "power",
            "voltage",
            "current",
            "load"
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

                IList<Element> mepElements = GetAllMepElements(doc);
                if (mepElements.Count == 0)
                {
                    TaskDialog.Show("Revit", "No MEP elements found in this model.");
                    return Result.Succeeded;
                }

                string csvPath = GetScheduleFilePath("MEP_TakeOff.csv");
                ExportElementsToCsv(doc, mepElements, csvPath);

                TaskDialog.Show(
                    "Revit Export",
                    $"Exported {mepElements.Count} MEP elements to:\n{csvPath}"
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

        private IList<Element> GetAllMepElements(Document doc)
        {
            List<BuiltInCategory> categories = new List<BuiltInCategory>
            {
                BuiltInCategory.OST_DuctCurves,
                BuiltInCategory.OST_DuctFitting,
                BuiltInCategory.OST_DuctAccessory,
                BuiltInCategory.OST_DuctTerminal,
                BuiltInCategory.OST_FlexDuctCurves,
                BuiltInCategory.OST_PipeCurves,
                BuiltInCategory.OST_PipeFitting,
                BuiltInCategory.OST_PipeAccessory,
                BuiltInCategory.OST_FlexPipeCurves,
                BuiltInCategory.OST_CableTray,
                BuiltInCategory.OST_CableTrayFitting,
                BuiltInCategory.OST_Conduit,
                BuiltInCategory.OST_ConduitFitting,
                BuiltInCategory.OST_PlumbingFixtures,
                BuiltInCategory.OST_MechanicalEquipment,
                BuiltInCategory.OST_ElectricalEquipment,
                BuiltInCategory.OST_ElectricalFixtures,
                BuiltInCategory.OST_LightingFixtures,
                BuiltInCategory.OST_Sprinklers
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

        private string GetScheduleFilePath(string fileName)
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
            return Path.Combine(schedulesDirectory, fileName);
        }

        private void ExportElementsToCsv(Document doc, IList<Element> elementsToExport, string filePath)
        {
            StringBuilder csv = new StringBuilder();
            csv.AppendLine(
                "ElementId,Category,Family,Type,Level,Mark,System Name,System Type,Service Type,Classification,Size,Diameter,Width,Height,Length,Area,Volume,Material,Weight,Unit Weight,Insulation Thickness,Lining Thickness,Airflow,Flow,Pressure Drop,Cooling Capacity,Heating Capacity,Power,Voltage,Current,Apparent Load,Connected Load,Connector Count,Connector Flow,Connector Demand,Connector Max Diameter (in),Connector Max Width (in),Connector Max Height (in),Comments,Parameter Snapshot"
            );

            foreach (Element elem in elementsToExport)
            {
                ConnectorMetrics connectorMetrics = GetConnectorMetrics(elem);

                string elementId = elem.Id.Value.ToString();
                string category = elem.Category?.Name ?? "";
                string family = GetFamilyName(elem);
                string typeName = GetTypeName(doc, elem);
                string level = GetLevelName(doc, elem);
                string mark = GetFirstAvailableParameterValue(doc, elem, "Mark");
                string systemName = GetFirstAvailableParameterValue(doc, elem, "System Name", "System");
                string systemType = GetFirstAvailableParameterValue(doc, elem, "System Type");
                string serviceType = GetFirstAvailableParameterValue(doc, elem, "Service Type");
                string classification = GetFirstAvailableParameterValue(
                    doc,
                    elem,
                    "Classification",
                    "Flow Classification",
                    "Part Type"
                );
                string size = GetFirstAvailableParameterValue(doc, elem, "Size", "Nominal Size", "Overall Size");
                string diameter = GetFirstAvailableParameterValue(
                    doc,
                    elem,
                    "Diameter",
                    "Nominal Diameter",
                    "Duct Diameter"
                );
                string width = GetFirstAvailableParameterValue(doc, elem, "Width", "Nominal Width", "Duct Width");
                string height = GetFirstAvailableParameterValue(doc, elem, "Height", "Nominal Height", "Duct Height");
                string length = GetFirstAvailableParameterValue(doc, elem, "Length", "Overall Size");
                string area = GetFirstAvailableParameterValue(doc, elem, "Area", "Surface Area");
                string volume = GetFirstAvailableParameterValue(doc, elem, "Volume");
                string material = GetMaterialSummary(doc, elem);
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
                    "Unit Weight",
                    "Weight per Unit Length",
                    "Mass per Unit Length"
                );
                string insulationThickness = GetFirstAvailableParameterValue(
                    doc,
                    elem,
                    "Insulation Thickness"
                );
                string liningThickness = GetFirstAvailableParameterValue(doc, elem, "Lining Thickness");
                string airflow = GetFirstAvailableParameterValue(
                    doc,
                    elem,
                    "Air Flow",
                    "Airflow",
                    "Calculated Supply Air Flow",
                    "Calculated Exhaust Air Flow",
                    "Calculated Return Air Flow",
                    "Flow"
                );
                string flow = GetFirstAvailableParameterValue(
                    doc,
                    elem,
                    "Flow",
                    "Flow Rate",
                    "Actual Flow",
                    "Demand Flow"
                );
                string pressureDrop = GetFirstAvailableParameterValue(
                    doc,
                    elem,
                    "Pressure Drop",
                    "Calculated Pressure Drop",
                    "Fitting Pressure Drop",
                    "Loss Method"
                );
                string coolingCapacity = GetFirstAvailableParameterValue(
                    doc,
                    elem,
                    "Cooling Capacity",
                    "Total Cooling Capacity",
                    "Sensible Cooling Capacity"
                );
                string heatingCapacity = GetFirstAvailableParameterValue(
                    doc,
                    elem,
                    "Heating Capacity",
                    "Heating Load",
                    "Total Heating Capacity"
                );
                string power = GetFirstAvailableParameterValue(
                    doc,
                    elem,
                    "Power",
                    "Power Factor",
                    "Motor Power",
                    "Input Power"
                );
                string voltage = GetFirstAvailableParameterValue(doc, elem, "Voltage");
                string current = GetFirstAvailableParameterValue(doc, elem, "Current", "Current Rating");
                string apparentLoad = GetFirstAvailableParameterValue(doc, elem, "Apparent Load");
                string connectedLoad = GetFirstAvailableParameterValue(doc, elem, "Connected Load", "Load Name");
                string comments = GetFirstAvailableParameterValue(doc, elem, "Comments");
                string parameterSnapshot = BuildParameterSnapshot(doc, elem);

                csv.AppendLine(string.Join(",",
                    EscapeCsv(elementId),
                    EscapeCsv(category),
                    EscapeCsv(family),
                    EscapeCsv(typeName),
                    EscapeCsv(level),
                    EscapeCsv(mark),
                    EscapeCsv(systemName),
                    EscapeCsv(systemType),
                    EscapeCsv(serviceType),
                    EscapeCsv(classification),
                    EscapeCsv(size),
                    EscapeCsv(diameter),
                    EscapeCsv(width),
                    EscapeCsv(height),
                    EscapeCsv(length),
                    EscapeCsv(area),
                    EscapeCsv(volume),
                    EscapeCsv(material),
                    EscapeCsv(weight),
                    EscapeCsv(unitWeight),
                    EscapeCsv(insulationThickness),
                    EscapeCsv(liningThickness),
                    EscapeCsv(airflow),
                    EscapeCsv(flow),
                    EscapeCsv(pressureDrop),
                    EscapeCsv(coolingCapacity),
                    EscapeCsv(heatingCapacity),
                    EscapeCsv(power),
                    EscapeCsv(voltage),
                    EscapeCsv(current),
                    EscapeCsv(apparentLoad),
                    EscapeCsv(connectedLoad),
                    EscapeCsv(connectorMetrics.Count.ToString(CultureInfo.InvariantCulture)),
                    EscapeCsv(FormatDouble(connectorMetrics.Flow)),
                    EscapeCsv(FormatDouble(connectorMetrics.Demand)),
                    EscapeCsv(FormatDouble(connectorMetrics.MaxDiameterInches)),
                    EscapeCsv(FormatDouble(connectorMetrics.MaxWidthInches)),
                    EscapeCsv(FormatDouble(connectorMetrics.MaxHeightInches)),
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

        private ConnectorMetrics GetConnectorMetrics(Element elem)
        {
            List<Connector> connectors = GetConnectors(elem);
            ConnectorMetrics metrics = new ConnectorMetrics();

            foreach (Connector connector in connectors)
            {
                metrics.Count += 1;
                metrics.Flow += SafeGetDouble(() => connector.Flow);
                metrics.Demand += SafeGetDouble(() => connector.Demand);
                metrics.MaxDiameterInches = Math.Max(
                    metrics.MaxDiameterInches,
                    SafeGetDouble(() => connector.Radius) * 24.0
                );
                metrics.MaxWidthInches = Math.Max(
                    metrics.MaxWidthInches,
                    SafeGetDouble(() => connector.Width) * 12.0
                );
                metrics.MaxHeightInches = Math.Max(
                    metrics.MaxHeightInches,
                    SafeGetDouble(() => connector.Height) * 12.0
                );
            }

            return metrics;
        }

        private List<Connector> GetConnectors(Element elem)
        {
            List<Connector> connectors = new List<Connector>();
            ConnectorSet? connectorSet = null;

            if (elem is MEPCurve mepCurve)
            {
                connectorSet = mepCurve.ConnectorManager?.Connectors;
            }
            else if (elem is FamilyInstance familyInstance && familyInstance.MEPModel != null)
            {
                connectorSet = familyInstance.MEPModel.ConnectorManager?.Connectors;
            }

            if (connectorSet == null)
                return connectors;

            foreach (Connector connector in connectorSet)
            {
                connectors.Add(connector);
            }

            return connectors;
        }

        private double SafeGetDouble(Func<double> getter)
        {
            try
            {
                return getter();
            }
            catch
            {
                return 0.0;
            }
        }

        private string FormatDouble(double value)
        {
            return Math.Abs(value) < 1e-9
                ? ""
                : value.ToString("0.###", CultureInfo.InvariantCulture);
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

        private class ConnectorMetrics
        {
            public int Count { get; set; }
            public double Flow { get; set; }
            public double Demand { get; set; }
            public double MaxDiameterInches { get; set; }
            public double MaxWidthInches { get; set; }
            public double MaxHeightInches { get; set; }
        }
    }
}
