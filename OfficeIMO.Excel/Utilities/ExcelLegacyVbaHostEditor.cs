using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Excel.Utilities;

/// <summary>Maintains Excel worksheet document modules through the shared, source-only VBA owner.</summary>
internal static class ExcelLegacyVbaHostEditor {
    private static readonly Guid WorksheetClass = new("00020820-0000-0000-C000-000000000046");

    internal static byte[] Reconcile(byte[] bytes, IReadOnlyList<ExcelLegacySheetBindings.Binding> bindings,
        IReadOnlyDictionary<string, string?> originalWorksheetCodeNames, Dictionary<string, string> allocatedNames,
        IDictionary<string, string> newCodeNames) {
        var removed = originalWorksheetCodeNames.Where(item => !bindings.Any(binding => binding.Original?.Name == item.Key))
            .Where(item => !string.IsNullOrEmpty(item.Value)).ToArray();
        var added = bindings.Where(binding => binding.Original == null && binding.Part is WorksheetPart).ToArray();
        if (removed.Length == 0 && added.Length == 0) return bytes;
        var project = OfficeVbaProject.Load(bytes);
        foreach (var item in removed) {
            var module = project.Modules.FirstOrDefault(module => module.Name.Equals(item.Value, StringComparison.OrdinalIgnoreCase));
            if (module != null) project.DeleteDocumentModule(module.Name);
        }
        var used = new HashSet<string>(project.Modules.Select(module => module.Name), StringComparer.OrdinalIgnoreCase);
        foreach (var binding in added) {
            string uri = binding.Part.Uri.OriginalString;
            string? name = ((WorksheetPart)binding.Part).Worksheet?.SheetProperties?.CodeName?.Value;
            if (string.IsNullOrEmpty(name) && allocatedNames.TryGetValue(uri, out string? allocated)) name = allocated;
            if (string.IsNullOrEmpty(name)) {
                int index = 1;
                do { name = "Sheet" + (index++).ToString(System.Globalization.CultureInfo.InvariantCulture); } while (used.Contains(name));
            }
            if (!used.Add(name!)) throw new NotSupportedException("The new worksheet code name collides with an existing VBA module.");
            project.AddDocumentModule(name!, string.Empty, WorksheetClass);
            allocatedNames[uri] = name!; newCodeNames[binding.Name] = name!;
        }
        return project.Write().GetBytes();
    }
}
