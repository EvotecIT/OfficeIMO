using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Utilities;

namespace OfficeIMO.Excel.LegacyXls.Write;

internal sealed partial class LegacyXlsPreservationSnapshot {
    private static OpenXmlElement NormalizeReferences(OpenXmlElement root, IReadOnlyDictionary<string, string>? names) =>
        ExcelPackageReferenceNormalizer.Copy(root, names);

    private bool ValidateDefinedNames(ExcelDocument document, IReadOnlyList<ExcelLegacySheetBindings.Binding> bindings,
        IReadOnlyDictionary<string, string> renames) {
        var expected = (DefinedNames?)_definedNames?.CloneNode(true);
        if (expected != null) foreach (var name in expected.Elements<DefinedName>().ToArray()) {
            if (name.LocalSheetId?.Value is not uint oldIndex) continue;
            var current = bindings.FirstOrDefault(binding => binding.Original?.ProjectedIndex == oldIndex);
            if (current == null) name.Remove(); else name.LocalSheetId = checked((uint)current.Index);
        }
        var actual = document.WorkbookRoot.DefinedNames == null ? null
            : (DefinedNames)NormalizeReferences(document.WorkbookRoot.DefinedNames, renames);
        string Canonical(DefinedNames? names) => names?.HasChildren == true ? names.OuterXml : string.Empty;
        return Canonical(expected) == Canonical(actual);
    }

    private bool PartsMatch(Dictionary<string, string> actual, IReadOnlyList<ExcelLegacySheetBindings.Binding> bindings) {
        var currentUris = new HashSet<string>(bindings.Select(binding => binding.Part.Uri.OriginalString), StringComparer.Ordinal);
        foreach (var original in _bindings.Originals)
            if (!original.Value.Worksheet && !currentUris.Contains(original.Key)) return false;
        foreach (var item in _parts) {
            if (actual.TryGetValue(item.Key, out string? value)) { if (value != item.Value) return false; continue; }
            if (!_partOwners.TryGetValue(item.Key, out var owners) || owners.Any(currentUris.Contains)) return false;
        }
        var newWorksheets = new HashSet<string>(bindings.Where(binding => binding.Original == null && binding.Part is WorksheetPart)
            .Select(binding => binding.Part.Uri.OriginalString), StringComparer.Ordinal);
        return actual.Keys.All(key => _parts.ContainsKey(key) || newWorksheets.Contains(key));
    }

    private static Dictionary<string, HashSet<string>> CapturePartOwners(ExcelDocument document) => ExcelWorksheetPartOwnership.Capture(document);
}
