using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel.Utilities;

/// <summary>Tracks native sheet identity through projected package parts rather than mutable tab names or positions.</summary>
internal sealed class ExcelLegacySheetBindings {
    private readonly Dictionary<string, OriginalSheet> _original;
    internal ExcelLegacySheetBindings(ExcelDocument document, IReadOnlyDictionary<string, int> nativeIndexes) {
        _original = Read(document).ToDictionary(binding => binding.Part.Uri.OriginalString,
            binding => new OriginalSheet(binding.Name, nativeIndexes[binding.Name], binding.Part is WorksheetPart,
                binding.Part is WorksheetPart worksheet ? worksheet.Worksheet?.SheetProperties?.CodeName?.Value : null, binding.Index), StringComparer.Ordinal);
    }

    internal IReadOnlyList<Binding> ReadCurrent(ExcelDocument document) => Read(document).Select(binding => {
        _original.TryGetValue(binding.Part.Uri.OriginalString, out var original);
        return binding with { Original = original };
    }).ToArray();

    internal IReadOnlyDictionary<string, string> GetReverseRenames(ExcelDocument document) => ReadCurrent(document)
        .Where(binding => binding.Original != null && binding.Name != binding.Original.Name)
        .ToDictionary(binding => binding.Name, binding => binding.Original!.Name, StringComparer.OrdinalIgnoreCase);

    internal IReadOnlyDictionary<string, OriginalSheet> Originals => _original;
    internal IReadOnlyCollection<string> OriginalUris => _original.Keys;

    private static IReadOnlyList<Binding> Read(ExcelDocument document) {
        var result = new List<Binding>(); var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (var sheet in document.WorkbookRoot.Sheets?.Elements<Sheet>() ?? Enumerable.Empty<Sheet>()) {
            string id = sheet.Id?.Value ?? throw new NotSupportedException("A native sheet requires a package relationship.");
            var part = document.WorkbookPartRoot.GetPartById(id);
            if (part is not WorksheetPart && part is not ChartsheetPart) throw new NotSupportedException("The native sheet type is not supported for structure editing.");
            if (!seen.Add(part.Uri.OriginalString)) throw new NotSupportedException("A worksheet part cannot have multiple native sheet identities.");
            result.Add(new Binding(sheet.Name?.Value ?? string.Empty, sheet, part, result.Count, null));
        }
        return result;
    }

    internal sealed record OriginalSheet(string Name, int NativeIndex, bool Worksheet, string? CodeName, int ProjectedIndex);
    internal sealed record Binding(string Name, Sheet Element, OpenXmlPart Part, int Index, OriginalSheet? Original);
}
