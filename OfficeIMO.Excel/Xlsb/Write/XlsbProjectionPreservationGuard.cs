using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel.Utilities;

namespace OfficeIMO.Excel.Xlsb.Write;

/// <summary>Protects retained chart/VBA graphs and detects direct worksheet DOM edits before unchanged-source copying.</summary>
internal sealed class XlsbProjectionPreservationGuard {
    private readonly Dictionary<string, string> _preservedParts;
    private readonly Dictionary<string, string> _worksheets;
    private readonly ExcelLegacySheetBindings _bindings;
    private readonly Dictionary<string, HashSet<string>> _owners;
    internal Dictionary<string, string> AllocatedCodeNames { get; } = new(StringComparer.Ordinal);

    private XlsbProjectionPreservationGuard(ExcelDocument document) {
        _bindings = new ExcelLegacySheetBindings(document, document.WorkbookRoot.Sheets!.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>()
            .Select((sheet, index) => (sheet, index)).ToDictionary(item => item.sheet.Name!.Value!, item => item.index, StringComparer.Ordinal));
        _owners = ExcelWorksheetPartOwnership.Capture(document);
        _preservedParts = CaptureParts(document, worksheets: false);
        _worksheets = CaptureParts(document, worksheets: true);
    }

    internal static XlsbProjectionPreservationGuard Capture(ExcelDocument document) => new(document);

    internal void Validate(ExcelDocument document) {
        var bindings = ReadCurrent(document);
        var currentUris = new HashSet<string>(bindings.Select(binding => binding.Part.Uri.OriginalString), StringComparer.Ordinal);
        var actual = CaptureParts(document, worksheets: false, _bindings.GetReverseRenames(document));
        bool valid = _bindings.Originals.Where(item => !item.Value.Worksheet).All(item => currentUris.Contains(item.Key))
            && _preservedParts.All(item => actual.TryGetValue(item.Key, out string? value) ? value == item.Value
                : _owners.TryGetValue(item.Key, out var owners) && owners.All(owner => !currentUris.Contains(owner)))
            && actual.Keys.All(_preservedParts.ContainsKey);
        if (!valid) throw new NotSupportedException("Native XLSB rewriting cannot modify retained chart, VBA or related-part graphs.");
    }

    internal IReadOnlyList<ExcelLegacySheetBindings.Binding> ReadCurrent(ExcelDocument document) => _bindings.ReadCurrent(document);
    internal IReadOnlyDictionary<string, string> ReverseRenames(ExcelDocument document) => _bindings.GetReverseRenames(document);

    internal bool HasWorksheetChanges(ExcelDocument document) => !Matches(_worksheets, CaptureParts(document, worksheets: true));

    private static bool Matches(Dictionary<string, string> expected, Dictionary<string, string> actual) =>
        expected.Count == actual.Count && expected.All(item => actual.TryGetValue(item.Key, out string? value) && value == item.Value);

    private static Dictionary<string, string> CaptureParts(ExcelDocument document, bool worksheets, IReadOnlyDictionary<string, string>? renames = null) {
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        var visited = new HashSet<OpenXmlPart>(); var queue = new Queue<OpenXmlPart>();
        foreach (var pair in document.OpenXmlDocument.Parts) queue.Enqueue(pair.OpenXmlPart);
        while (queue.Count > 0) {
            OpenXmlPart part = queue.Dequeue();
            if (!visited.Add(part)) continue;
            foreach (var pair in part.Parts) queue.Enqueue(pair.OpenXmlPart);
            bool editablePart = part is WorksheetPart or WorkbookPart or WorkbookStylesPart or SharedStringTablePart;
            if (worksheets ? !editablePart : part is WorksheetPart or WorkbookStylesPart or SharedStringTablePart) continue;
            result.Add(part.Uri.OriginalString, ExcelPackagePartFingerprint.Compute(part, includeContent: worksheets || part is not WorkbookPart,
                writeContent: part.RootElement == null || renames == null || renames.Count == 0 ? null : writer => ExcelPackageReferenceNormalizer.Copy(part.RootElement, renames).WriteTo(writer),
                includeChildPart: part is WorkbookPart ? child => child is not WorksheetPart : null));
        }
        return result;
    }

}
