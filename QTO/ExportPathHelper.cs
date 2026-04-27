using System;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;

namespace QTO
{
    internal static class ExportPathHelper
    {
        public static string GetScheduleFilePath(Document doc, string exportSuffix)
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

            string modelBaseName = GetModelBaseName(doc);
            return Path.Combine(schedulesDirectory, $"{modelBaseName}_{exportSuffix}.csv");
        }

        private static string GetModelBaseName(Document doc)
        {
            string rawName = !string.IsNullOrWhiteSpace(doc.PathName)
                ? Path.GetFileNameWithoutExtension(doc.PathName)
                : doc.Title;

            if (string.IsNullOrWhiteSpace(rawName))
            {
                rawName = "Untitled_Model";
            }

            char[] invalidChars = Path.GetInvalidFileNameChars();
            string sanitized = new string(
                rawName
                    .Select(ch => invalidChars.Contains(ch) ? '_' : ch)
                    .ToArray()
            ).Trim();

            return string.IsNullOrWhiteSpace(sanitized) ? "Untitled_Model" : sanitized;
        }
    }
}
