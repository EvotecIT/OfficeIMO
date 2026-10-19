using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.LegacyXls.Model;
using OfficeIMO.Excel.Utilities;
using System.Xml;
using System.Threading;

namespace OfficeIMO.Excel.LegacyXls.Write;

/// <summary>Admits literal edits only when the projected workbook's other contracts remain unchanged.</summary>
internal sealed class LegacyXlsPreservationSnapshot {
    private readonly Dictionary<string, string> _parts;
    private readonly Dictionary<(string Sheet, string Cell), CellValueSnapshot> _cells;
    private readonly LegacyXlsPreservingWorkbook _source;
    private readonly string _customProperties;
    private readonly string _coreProperties;

    private LegacyXlsPreservationSnapshot(ExcelDocument document, LegacyXlsWorkbook workbook, CancellationToken cancellationToken) {
        _source = new LegacyXlsPreservingWorkbook(workbook, cancellationToken);
        _customProperties = CaptureCustomProperties(document);
        _coreProperties = document.GetLegacyXlsCorePropertiesSnapshot();
        _parts = CaptureParts(document, cancellationToken);
        _cells = CaptureCells(document, cancellationToken);
    }

    internal static LegacyXlsPreservationSnapshot? Capture(ExcelDocument document, LegacyXlsWorkbook workbook, CancellationToken cancellationToken) {
        if (workbook.SourceCompoundFile == null || workbook.WasEncryptedSource
            || workbook.CompoundFeatureRecords.Any(item => item.Kind == LegacyXlsCompoundFeatureRecordKind.DigitalSignature)) return null;
        try { return new LegacyXlsPreservationSnapshot(document, workbook, cancellationToken); }
        catch (NotSupportedException) { return null; }
    }

    internal bool TryCreatePlan(ExcelDocument document, out LegacyXlsPreservationPlan? plan, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        plan = null;
        if (document.GetLegacyXlsCorePropertiesSnapshot() != _coreProperties || CaptureCustomProperties(document) != _customProperties) return false;
        var currentParts = CaptureParts(document, cancellationToken);
        if (_parts.Count != currentParts.Count || _parts.Any(item => !currentParts.TryGetValue(item.Key, out string? value) || value != item.Value)) return false;
        var currentCells = CaptureCells(document, cancellationToken);
        if (_cells.Count != currentCells.Count) return false;
        var changes = new List<LegacyXlsLiteralEdit>();
        foreach (var item in _cells) {
            if (!currentCells.TryGetValue(item.Key, out var current)) return false;
            if (current.Formatting != item.Value.Formatting) return false;
            if (current.Text == item.Value.Text && current.Kind == item.Value.Kind) continue;
            if (current.Formatting.Length != 0) return false;
            if (item.Value.IsFormula || current.IsFormula) return false;
            var change = new LegacyXlsLiteralEdit(item.Key.Sheet, item.Key.Cell, current.Text, current.Kind);
            if (!_source.CanApply(change)) return false;
            changes.Add(change);
        }
        plan = new LegacyXlsPreservationPlan(_source, changes);
        return true;
    }

    // Compare values rather than a dirty flag: import can populate pending properties before this baseline.
    private static string CaptureCustomProperties(ExcelDocument document) => string.Join("\0",
        document.CustomDocumentProperties.OrderBy(item => item.Key, StringComparer.Ordinal)
            .Select(item => item.Value.ToOpenXml(item.Key).OuterXml + "\0" +
                (item.Value.PropertyType == ExcelCustomPropertyType.DateTime
                    ? Convert.ToDateTime(item.Value.Value, System.Globalization.CultureInfo.InvariantCulture)
                        .ToUniversalTime().ToString("O", System.Globalization.CultureInfo.InvariantCulture)
                    : string.Empty)));

    private static Dictionary<string, string> CaptureParts(ExcelDocument document, CancellationToken cancellationToken) {
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        var visited = new HashSet<OpenXmlPart>(); var queue = new Queue<OpenXmlPart>();
        foreach (var item in document.OpenXmlDocument.Parts) queue.Enqueue(item.OpenXmlPart);
        while (queue.Count != 0) {
            cancellationToken.ThrowIfCancellationRequested();
            var part = queue.Dequeue(); if (!visited.Add(part)) continue;
            foreach (var item in part.Parts) queue.Enqueue(item.OpenXmlPart);
            if (part is SharedStringTablePart) continue;
            result.Add(part.Uri.OriginalString, ExcelPackagePartFingerprint.Compute(part,
                writeContent: part is WorksheetPart worksheet ? writer => WriteWorksheetMetadata(writer, worksheet.Worksheet!, cancellationToken) : null));
        }
        return result;
    }

    private static void WriteWorksheetMetadata(XmlWriter writer, Worksheet worksheet, CancellationToken cancellationToken) {
        writer.WriteStartElement("worksheetSnapshot"); WriteAttributes(writer, worksheet, false);
        foreach (var child in worksheet.ChildElements) {
            if (child is not SheetData data) { child.WriteTo(writer); continue; }
            writer.WriteStartElement("data"); WriteAttributes(writer, data, false);
            foreach (var row in data.ChildElements) {
                cancellationToken.ThrowIfCancellationRequested();
                writer.WriteStartElement("row"); WriteAttributes(writer, row, false);
                foreach (var item in row.ChildElements) {
                    if (item is not Cell cell) { item.WriteTo(writer); continue; }
                    writer.WriteStartElement("cell"); WriteAttributes(writer, cell, true);
                    foreach (var element in cell.ChildElements)
                        if (element is not CellValue and not InlineString) element.WriteTo(writer);
                    writer.WriteEndElement();
                }
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
        }
        writer.WriteEndElement();
    }

    private static void WriteAttributes(XmlWriter writer, OpenXmlElement element, bool cell) {
        foreach (var attribute in element.GetAttributes().OrderBy(item => item.NamespaceUri, StringComparer.Ordinal).ThenBy(item => item.LocalName, StringComparer.Ordinal)) {
            if (cell && attribute.NamespaceUri.Length == 0 && attribute.LocalName == "t") continue;
            writer.WriteStartElement("attribute"); writer.WriteAttributeString("namespace", attribute.NamespaceUri);
            writer.WriteAttributeString("name", attribute.LocalName); writer.WriteAttributeString("value", attribute.Value); writer.WriteEndElement();
        }
    }

    private static Dictionary<(string, string), CellValueSnapshot> CaptureCells(ExcelDocument document, CancellationToken cancellationToken) {
        var cells = new Dictionary<(string, string), CellValueSnapshot>();
        using var sharedText = SharedStringCache.Build(document.OpenXmlDocument);
        var sharedItems = document.WorkbookPartRoot.SharedStringTablePart?.SharedStringTable?.Elements<SharedStringItem>().ToArray() ?? Array.Empty<SharedStringItem>();
        var formats = new Dictionary<int, string>();
        foreach (var sheet in document.Sheets)
            foreach (var cell in sheet.WorksheetPart.Worksheet!.Descendants<Cell>()) {
                cancellationToken.ThrowIfCancellationRequested();
                string address = cell.CellReference?.Value ?? throw new NotSupportedException("Native XLS preservation requires explicit cell addresses.");
                var type = cell.DataType?.Value;
                string kind = type == CellValues.SharedString || type == CellValues.InlineString || type == CellValues.String ? "text"
                    : type == CellValues.Boolean ? "boolean" : type == CellValues.Error ? "error" : type == CellValues.Date ? "date" : "number";
                cells.Add((sheet.Name, address), new CellValueSnapshot(type == CellValues.SharedString && int.TryParse(cell.CellValue?.Text, out int stringIndex) ? sharedText.Get(stringIndex) ?? string.Empty : sheet.GetCellText(cell), kind, cell.CellFormula != null, GetTextFormatting(cell, sharedItems, formats)));
            }
        return cells;
    }

    private static string GetTextFormatting(Cell cell, SharedStringItem[] sharedItems, Dictionary<int, string> formats) {
        OpenXmlElement? value = cell.InlineString;
        int? sharedIndex = null;
        if (cell.DataType?.Value == CellValues.SharedString && int.TryParse(cell.CellValue?.Text, out int index) && index >= 0 && index < sharedItems.Length) {
            if (formats.TryGetValue(index, out string? existing)) return existing;
            sharedIndex = index; value = sharedItems[index];
        }
        if (value == null) return string.Empty;
        var result = new System.Text.StringBuilder();
        foreach (var child in value.ChildElements.Where(item => item is not Text)) {
            var clone = child.CloneNode(true);
            foreach (var text in clone.Descendants<Text>().ToArray()) text.Remove();
            result.Append(clone.OuterXml);
        }
        string format = result.ToString();
        if (sharedIndex.HasValue) formats[sharedIndex.Value] = format;
        return format;
    }

    private sealed record CellValueSnapshot(string Text, string Kind, bool IsFormula, string Formatting);
}

internal sealed record LegacyXlsLiteralEdit(string Sheet, string Cell, string Text, string Kind);

internal sealed class LegacyXlsPreservationPlan {
    private readonly LegacyXlsPreservingWorkbook _source;
    private readonly IReadOnlyList<LegacyXlsLiteralEdit> _changes;
    internal LegacyXlsPreservationPlan(LegacyXlsPreservingWorkbook source, IReadOnlyList<LegacyXlsLiteralEdit> changes) { _source = source; _changes = changes; }
    internal byte[] Write(byte[]? filePass, long maximumBytes, System.Threading.CancellationToken cancellationToken) => _source.Write(_changes, filePass, maximumBytes, cancellationToken);
    internal byte[] WriteEncrypted(byte[] filePass, string password, byte[] salt, long maximumBytes) =>
        _source.Write(_changes, filePass, maximumBytes, CancellationToken.None, bytes => OfficeIMO.Excel.LegacyXls.Biff.BiffRc4Encryption.EncryptWorkbookStream(bytes, password, salt));
}
