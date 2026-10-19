using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.LegacyXls.Model;
using OfficeIMO.Excel.Utilities;
using System.Xml;
using System.Threading;

namespace OfficeIMO.Excel.LegacyXls.Write;

/// <summary>Admits literal edits only when the projected workbook's other contracts remain unchanged.</summary>
internal sealed partial class LegacyXlsPreservationSnapshot {
    private readonly Dictionary<string, string> _parts;
    private readonly Dictionary<(string Sheet, string Cell), CellValueSnapshot> _cells;
    private readonly LegacyXlsPreservingWorkbook _source;
    private readonly ExcelLegacySheetBindings _bindings;
    private readonly string _sheetBindings;
    private readonly DefinedNames? _definedNames;
    private readonly Dictionary<string, HashSet<string>> _partOwners;
    private readonly string _customProperties;
    private readonly string _coreProperties;

    private LegacyXlsPreservationSnapshot(ExcelDocument document, LegacyXlsWorkbook workbook, CancellationToken cancellationToken) {
        _source = new LegacyXlsPreservingWorkbook(workbook, cancellationToken);
        _bindings = new ExcelLegacySheetBindings(document, _source.NativeIndexes);
        _sheetBindings = document.WorkbookRoot.Sheets?.OuterXml ?? string.Empty;
        _definedNames = (DefinedNames?)document.WorkbookRoot.DefinedNames?.CloneNode(true);
        _partOwners = CapturePartOwners(document);
        _customProperties = CaptureCustomProperties(document);
        _coreProperties = document.GetLegacyXlsCorePropertiesSnapshot();
        _parts = CaptureParts(document, cancellationToken);
        _cells = CaptureCells(document, cancellationToken);
    }

    internal static LegacyXlsPreservationSnapshot? Capture(ExcelDocument document, LegacyXlsWorkbook workbook, CancellationToken cancellationToken) {
        if (workbook.SourceCompoundFile == null
            || workbook.CompoundFeatureRecords.Any(item => item.Kind == LegacyXlsCompoundFeatureRecordKind.DigitalSignature)) return null;
        try { return new LegacyXlsPreservationSnapshot(document, workbook, cancellationToken); }
        catch (NotSupportedException) { return null; }
    }

    internal bool TryCreatePlan(ExcelDocument document, out LegacyXlsPreservationPlan? plan, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        plan = null;
        if (document.GetLegacyXlsCorePropertiesSnapshot() != _coreProperties || CaptureCustomProperties(document) != _customProperties) return false;
        var bindings = _bindings.ReadCurrent(document);
        var renames = _bindings.GetReverseRenames(document);
        bool structureChanged = _sheetBindings != (document.WorkbookRoot.Sheets?.OuterXml ?? string.Empty);
        if (!ValidateDefinedNames(document, bindings, renames)) return false;
        var currentParts = CaptureParts(document, cancellationToken, renames);
        if (!PartsMatch(currentParts, bindings)) return false;
        var currentCells = CaptureCells(document, cancellationToken, renames);
        var changes = new List<LegacyXlsLiteralEdit>();
        var edits = new List<LegacyXlsCellEdit>();
        bool literalOnly = !structureChanged;
        var currentSheetUris = new HashSet<string>(bindings.Select(binding => binding.Part.Uri.OriginalString), StringComparer.Ordinal);
        foreach (var key in _cells.Keys.Union(currentCells.Keys)) {
            bool existed = _cells.TryGetValue(key, out var previous);
            bool exists = currentCells.TryGetValue(key, out var current);
            if (existed && exists && current!.Text == previous!.Text && current.Kind == previous.Kind && current.Metadata == previous.Metadata && current.Style == previous.Style && current.Formatting == previous.Formatting) continue;
            if (!exists && !currentSheetUris.Contains(key.Item1)) continue;
            string name = exists ? current!.SheetName : bindings.Single(binding => binding.Part.Uri.OriginalString == key.Item1).Name;
            edits.Add(new LegacyXlsCellEdit(name, key.Item2, exists ? (Cell)current!.Cell.CloneNode(true) : null));
            if (!existed || !exists || previous!.Metadata != current!.Metadata || previous.Style != current.Style || previous.IsFormula || current.IsFormula || current.Formatting.Length != 0) { literalOnly = false; continue; }
            var change = new LegacyXlsLiteralEdit(name, key.Item2, current.Text, current.Kind);
            if (!_source.CanApply(change)) literalOnly = false;
            changes.Add(change);
        }
        plan = literalOnly ? new LegacyXlsPreservationPlan(_source, changes)
            : new LegacyXlsPreservationPlan(_source, document, edits, structureChanged ? bindings : null);
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

    private static Dictionary<string, string> CaptureParts(ExcelDocument document, CancellationToken cancellationToken, IReadOnlyDictionary<string, string>? renames = null) {
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        var visited = new HashSet<OpenXmlPart>(); var queue = new Queue<OpenXmlPart>();
        foreach (var item in document.OpenXmlDocument.Parts) queue.Enqueue(item.OpenXmlPart);
        while (queue.Count != 0) {
            cancellationToken.ThrowIfCancellationRequested();
            var part = queue.Dequeue(); if (!visited.Add(part)) continue;
            foreach (var item in part.Parts) queue.Enqueue(item.OpenXmlPart);
            if (part is SharedStringTablePart or WorkbookStylesPart) continue;
            OpenXmlElement? root = part.RootElement;
            result.Add(part.Uri.OriginalString, ExcelPackagePartFingerprint.Compute(part,
                writeContent: root == null ? null : writer => {
                    if (root is Worksheet worksheet) WriteWorksheetMetadata(writer, worksheet, cancellationToken, renames);
                    else if (root is Workbook workbook) {
                        var copy = (Workbook)NormalizeReferences(workbook, renames);
                        copy.Sheets = null; copy.DefinedNames = null; copy.WriteTo(writer);
                    } else NormalizeReferences(root, renames).WriteTo(writer);
                }, includeChildPart: child => part is not WorkbookPart || child is not WorksheetPart));
        }
        return result;
    }

    private static void WriteWorksheetMetadata(XmlWriter writer, Worksheet worksheet, CancellationToken cancellationToken, IReadOnlyDictionary<string, string>? renames) {
        writer.WriteStartElement("worksheetSnapshot"); WriteAttributes(writer, worksheet, false);
        foreach (var child in worksheet.ChildElements) {
            if (child is not SheetData data) { NormalizeReferences(child, renames).WriteTo(writer); continue; }
            writer.WriteStartElement("data"); WriteAttributes(writer, data, false);
            foreach (var row in data.ChildElements) {
                cancellationToken.ThrowIfCancellationRequested();
                if (row is not Row typedRow) { row.WriteTo(writer); continue; }
                var metadata = (Row)typedRow.CloneNode(true);
                metadata.RemoveAllChildren<Cell>(); metadata.RowIndex = null; metadata.Spans = null;
                if (metadata.HasAttributes || metadata.HasChildren) { metadata.RowIndex = typedRow.RowIndex?.Value; metadata.WriteTo(writer); }
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

    private static Dictionary<(string, string), CellValueSnapshot> CaptureCells(ExcelDocument document, CancellationToken cancellationToken, IReadOnlyDictionary<string, string>? renames = null) {
        var cells = new Dictionary<(string, string), CellValueSnapshot>();
        using var sharedText = SharedStringCache.Build(document.OpenXmlDocument);
        var sharedItems = document.WorkbookPartRoot.SharedStringTablePart?.SharedStringTable?.Elements<SharedStringItem>().ToArray() ?? Array.Empty<SharedStringItem>();
        var formats = new Dictionary<int, string>();
        var styles = new Dictionary<uint, string>();
        string Style(uint index) { if (!styles.TryGetValue(index, out var value)) styles.Add(index, value = CaptureStyle(document, index)); return value; }
        foreach (var sheet in document.Sheets)
            foreach (var cell in sheet.WorksheetPart.Worksheet!.Descendants<Cell>()) {
                cancellationToken.ThrowIfCancellationRequested();
                string address = cell.CellReference?.Value ?? throw new NotSupportedException("Native XLS preservation requires explicit cell addresses.");
                var type = cell.DataType?.Value;
                string kind = type == CellValues.SharedString || type == CellValues.InlineString || type == CellValues.String ? "text"
                    : type == CellValues.Boolean ? "boolean" : type == CellValues.Error ? "error" : type == CellValues.Date ? "date" : "number";
                cells.Add((sheet.WorksheetPart.Uri.OriginalString, address), new CellValueSnapshot(type == CellValues.SharedString && int.TryParse(cell.CellValue?.Text, out int stringIndex) ? sharedText.Get(stringIndex) ?? string.Empty : sheet.GetCellText(cell), kind, cell.CellFormula != null, GetTextFormatting(cell, sharedItems, formats), CaptureCellMetadata((Cell)NormalizeReferences(cell, renames)), Style(cell.StyleIndex?.Value ?? 0), cell, sheet.Name));
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

    private static string CaptureStyle(ExcelDocument document, uint index) {
        var sheet = document.WorkbookPartRoot.WorkbookStylesPart?.Stylesheet;
        var xf = sheet?.CellFormats?.Elements<CellFormat>().ElementAtOrDefault(checked((int)index));
        if (xf == null) return string.Empty;
        var parent = sheet?.CellStyleFormats?.Elements<CellFormat>().ElementAtOrDefault(checked((int)(xf.FormatId?.Value ?? 0)));
        string Components(CellFormat format) => string.Concat(
            sheet?.Fonts?.Elements<Font>().ElementAtOrDefault(checked((int)(format.FontId?.Value ?? 0)))?.OuterXml,
            sheet?.Fills?.Elements<Fill>().ElementAtOrDefault(checked((int)(format.FillId?.Value ?? 0)))?.OuterXml,
            sheet?.Borders?.Elements<Border>().ElementAtOrDefault(checked((int)(format.BorderId?.Value ?? 0)))?.OuterXml,
            sheet?.NumberingFormats?.Elements<NumberingFormat>().FirstOrDefault(item => item.NumberFormatId?.Value == format.NumberFormatId?.Value)?.OuterXml);
        return xf.OuterXml + Components(xf) + parent?.OuterXml + (parent == null ? string.Empty : Components(parent));
    }

    private static string CaptureCellMetadata(Cell cell) {
        var clone = (Cell)cell.CloneNode(true); clone.RemoveAllChildren<CellValue>(); clone.RemoveAllChildren<InlineString>(); clone.DataType = null;
        return clone.OuterXml;
    }

    private sealed record CellValueSnapshot(string Text, string Kind, bool IsFormula, string Formatting, string Metadata, string Style, Cell Cell, string SheetName);
}

internal sealed record LegacyXlsLiteralEdit(string Sheet, string Cell, string Text, string Kind);

internal sealed class LegacyXlsPreservationPlan {
    private readonly LegacyXlsPreservingWorkbook _source;
    private readonly IReadOnlyList<LegacyXlsLiteralEdit> _changes = Array.Empty<LegacyXlsLiteralEdit>();
    private readonly ExcelDocument? _document;
    private readonly IReadOnlyList<LegacyXlsCellEdit>? _edits;
    private readonly IReadOnlyList<ExcelLegacySheetBindings.Binding>? _bindings;
    internal LegacyXlsPreservationPlan(LegacyXlsPreservingWorkbook source, IReadOnlyList<LegacyXlsLiteralEdit> changes) { _source = source; _changes = changes; }
    internal LegacyXlsPreservationPlan(LegacyXlsPreservingWorkbook source, ExcelDocument document, IReadOnlyList<LegacyXlsCellEdit> edits, IReadOnlyList<ExcelLegacySheetBindings.Binding>? bindings = null) { _source = source; _document = document; _edits = edits; _bindings = bindings; }
    internal byte[] Write(byte[]? filePass, long maximumBytes, System.Threading.CancellationToken cancellationToken) => _edits == null ? _source.Write(_changes, filePass, maximumBytes, cancellationToken) : _source.WriteEdits(_document!, _edits, filePass, maximumBytes, cancellationToken, bindings: _bindings);
    internal byte[] WriteEncrypted(byte[] filePass, string password, byte[] salt, long maximumBytes) =>
        _edits == null ? _source.Write(_changes, filePass, maximumBytes, CancellationToken.None, bytes => OfficeIMO.Excel.LegacyXls.Biff.BiffRc4Encryption.EncryptWorkbookStream(bytes, password, salt))
        : _source.WriteEdits(_document!, _edits, filePass, maximumBytes, CancellationToken.None, bytes => OfficeIMO.Excel.LegacyXls.Biff.BiffRc4Encryption.EncryptWorkbookStream(bytes, password, salt), _bindings);
}
