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
    public class Push_TaskName_To_Revit : IExternalCommand
    {
        private const string ParameterName = "4D_Build_Code";
        private const string MappingFileName = "Revit_4D_Build_Code_Map.csv";

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            try
            {
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;

                string mappingPath = GetMappingFilePath();
                if (!File.Exists(mappingPath))
                {
                    TaskDialog.Show(
                        "Push 4D Build Codes",
                        $"Could not find the mapping file:\n{mappingPath}\n\nRun the planning pipeline first so {MappingFileName} exists."
                    );
                    return Result.Cancelled;
                }

                MappingLoadResult mappingResult = LoadMappings(mappingPath);
                if (mappingResult.Assignments.Count == 0)
                {
                    TaskDialog.Show(
                        "Push 4D Build Codes",
                        $"No non-empty 4D build code assignments were found in:\n{mappingPath}"
                    );
                    return Result.Cancelled;
                }

                WriteResult writeResult = WriteBuildCodesToElements(doc, mappingResult.Assignments);

                TaskDialog.Show("Push 4D Build Codes", BuildSummary(mappingPath, mappingResult, writeResult));
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("Error", ex.ToString());
                return Result.Failed;
            }
        }

        private string GetMappingFilePath()
        {
            string assemblyDirectory = Path.GetDirectoryName(
                System.Reflection.Assembly.GetExecutingAssembly().Location
            ) ?? "";

            DirectoryInfo? directory = new DirectoryInfo(assemblyDirectory);
            while (directory != null &&
                   !Directory.Exists(Path.Combine(directory.FullName, "src", "Planning_engine")))
            {
                directory = directory.Parent;
            }

            string planningDirectory = directory != null
                ? Path.Combine(directory.FullName, "src", "Planning_engine")
                : Path.Combine(assemblyDirectory, "src", "Planning_engine");

            return Path.Combine(planningDirectory, "Fuzor_Mapper", "outputs", MappingFileName);
        }

        private MappingLoadResult LoadMappings(string mappingPath)
        {
            Dictionary<int, HashSet<string>> buildCodesByElementId = new Dictionary<int, HashSet<string>>();
            int invalidElementIdCount = 0;
            int blankBuildCodeCount = 0;

            using StreamReader reader = new StreamReader(mappingPath);
            string? headerLine = reader.ReadLine();
            if (string.IsNullOrWhiteSpace(headerLine))
            {
                throw new InvalidOperationException($"The mapping file is empty: {mappingPath}");
            }

            List<string> headers = ParseCsvLine(headerLine);
            int elementIdIndex = headers.FindIndex(header => header.Equals("element_id", StringComparison.OrdinalIgnoreCase));
            int buildCodeIndex = headers.FindIndex(header => header.Equals("build_code", StringComparison.OrdinalIgnoreCase));

            if (elementIdIndex < 0 || buildCodeIndex < 0)
            {
                throw new InvalidOperationException(
                    $"The mapping file must include 'element_id' and 'build_code' columns: {mappingPath}"
                );
            }

            while (!reader.EndOfStream)
            {
                string? line = reader.ReadLine();
                if (string.IsNullOrWhiteSpace(line))
                {
                    continue;
                }

                List<string> fields = ParseCsvLine(line);
                if (fields.Count <= Math.Max(elementIdIndex, buildCodeIndex))
                {
                    continue;
                }

                string elementIdText = fields[elementIdIndex].Trim();
                string buildCode = fields[buildCodeIndex].Trim();

                if (!int.TryParse(elementIdText, out int elementId))
                {
                    invalidElementIdCount++;
                    continue;
                }

                if (string.IsNullOrWhiteSpace(buildCode))
                {
                    blankBuildCodeCount++;
                    continue;
                }

                if (!buildCodesByElementId.TryGetValue(elementId, out HashSet<string>? buildCodes))
                {
                    buildCodes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    buildCodesByElementId[elementId] = buildCodes;
                }

                buildCodes.Add(buildCode);
            }

            Dictionary<int, string> assignments = new Dictionary<int, string>();
            int conflictingBuildCodeCount = 0;

            foreach (KeyValuePair<int, HashSet<string>> pair in buildCodesByElementId)
            {
                if (pair.Value.Count == 1)
                {
                    assignments[pair.Key] = pair.Value.First();
                }
                else if (pair.Value.Count > 1)
                {
                    conflictingBuildCodeCount++;
                }
            }

            return new MappingLoadResult(
                assignments,
                invalidElementIdCount,
                blankBuildCodeCount,
                conflictingBuildCodeCount
            );
        }

        private WriteResult WriteBuildCodesToElements(Document doc, IReadOnlyDictionary<int, string> assignments)
        {
            int updatedCount = 0;
            int unchangedCount = 0;
            int missingElementCount = 0;
            int missingParameterCount = 0;
            int readOnlyParameterCount = 0;
            int incompatibleParameterCount = 0;

            using Transaction transaction = new Transaction(doc, "Push 4D build codes to Revit");
            transaction.Start();

            foreach (KeyValuePair<int, string> assignment in assignments)
            {
                Element? element = doc.GetElement(new ElementId(assignment.Key));
                if (element == null)
                {
                    missingElementCount++;
                    continue;
                }

                Parameter? parameter = element.LookupParameter(ParameterName);
                if (parameter == null)
                {
                    missingParameterCount++;
                    continue;
                }

                if (parameter.IsReadOnly)
                {
                    readOnlyParameterCount++;
                    continue;
                }

                if (parameter.StorageType != StorageType.String)
                {
                    incompatibleParameterCount++;
                    continue;
                }

                string currentValue = parameter.AsString() ?? string.Empty;
                if (string.Equals(currentValue, assignment.Value, StringComparison.Ordinal))
                {
                    unchangedCount++;
                    continue;
                }

                parameter.Set(assignment.Value);
                updatedCount++;
            }

            transaction.Commit();

            return new WriteResult(
                updatedCount,
                unchangedCount,
                missingElementCount,
                missingParameterCount,
                readOnlyParameterCount,
                incompatibleParameterCount
            );
        }

        private string BuildSummary(string mappingPath, MappingLoadResult mappingResult, WriteResult writeResult)
        {
            StringBuilder summary = new StringBuilder();
            summary.AppendLine($"Source: {mappingPath}");
            summary.AppendLine($"Target parameter: {ParameterName}");
            summary.AppendLine();
            summary.AppendLine($"Assignments loaded: {mappingResult.Assignments.Count}");
            summary.AppendLine($"Elements updated: {writeResult.UpdatedCount}");
            summary.AppendLine($"Already up to date: {writeResult.UnchangedCount}");
            summary.AppendLine($"Elements not found in model: {writeResult.MissingElementCount}");
            summary.AppendLine($"Elements missing '{ParameterName}': {writeResult.MissingParameterCount}");
            summary.AppendLine($"Read-only '{ParameterName}': {writeResult.ReadOnlyParameterCount}");
            summary.AppendLine($"Non-text '{ParameterName}': {writeResult.IncompatibleParameterCount}");

            if (mappingResult.BlankBuildCodeCount > 0)
            {
                summary.AppendLine($"Skipped blank build codes: {mappingResult.BlankBuildCodeCount}");
            }

            if (mappingResult.InvalidElementIdCount > 0)
            {
                summary.AppendLine($"Skipped invalid element ids: {mappingResult.InvalidElementIdCount}");
            }

            if (mappingResult.ConflictingBuildCodeCount > 0)
            {
                summary.AppendLine($"Skipped elements with conflicting build codes: {mappingResult.ConflictingBuildCodeCount}");
            }

            if (writeResult.MissingParameterCount > 0 || writeResult.ReadOnlyParameterCount > 0 || writeResult.IncompatibleParameterCount > 0)
            {
                summary.AppendLine();
                summary.AppendLine($"Make sure '{ParameterName}' exists on the target categories as an editable text parameter.");
            }

            return summary.ToString().TrimEnd();
        }

        private List<string> ParseCsvLine(string line)
        {
            List<string> fields = new List<string>();
            StringBuilder current = new StringBuilder();
            bool inQuotes = false;

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];

                if (c == '"')
                {
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        current.Append('"');
                        i++;
                    }
                    else
                    {
                        inQuotes = !inQuotes;
                    }

                    continue;
                }

                if (c == ',' && !inQuotes)
                {
                    fields.Add(current.ToString());
                    current.Clear();
                    continue;
                }

                current.Append(c);
            }

            fields.Add(current.ToString());
            return fields;
        }

        private sealed class MappingLoadResult
        {
            public MappingLoadResult(
                Dictionary<int, string> assignments,
                int invalidElementIdCount,
                int blankBuildCodeCount,
                int conflictingBuildCodeCount)
            {
                Assignments = assignments;
                InvalidElementIdCount = invalidElementIdCount;
                BlankBuildCodeCount = blankBuildCodeCount;
                ConflictingBuildCodeCount = conflictingBuildCodeCount;
            }

            public Dictionary<int, string> Assignments { get; }

            public int InvalidElementIdCount { get; }

            public int BlankBuildCodeCount { get; }

            public int ConflictingBuildCodeCount { get; }
        }

        private sealed class WriteResult
        {
            public WriteResult(
                int updatedCount,
                int unchangedCount,
                int missingElementCount,
                int missingParameterCount,
                int readOnlyParameterCount,
                int incompatibleParameterCount)
            {
                UpdatedCount = updatedCount;
                UnchangedCount = unchangedCount;
                MissingElementCount = missingElementCount;
                MissingParameterCount = missingParameterCount;
                ReadOnlyParameterCount = readOnlyParameterCount;
                IncompatibleParameterCount = incompatibleParameterCount;
            }

            public int UpdatedCount { get; }

            public int UnchangedCount { get; }

            public int MissingElementCount { get; }

            public int MissingParameterCount { get; }

            public int ReadOnlyParameterCount { get; }

            public int IncompatibleParameterCount { get; }
        }
    }
}
