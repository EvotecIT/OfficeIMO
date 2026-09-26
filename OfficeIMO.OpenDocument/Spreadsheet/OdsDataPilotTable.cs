using OfficeIMO.Spreadsheet;

namespace OfficeIMO.OpenDocument;

/// <summary>An ODF data pilot table. Imported settings outside the typed field subset remain in the source XML.</summary>
public sealed class OdsDataPilotTable {
    private readonly OdsDocument _document;
    internal XElement Element { get; }

    internal OdsDataPilotTable(OdsDocument document, XElement element) {
        _document = document;
        Element = element;
    }

    /// <summary>Gets the data pilot name.</summary>
    public string Name => (string?)Element.Attribute(OdfNamespaces.Table + "name") ?? string.Empty;

    /// <summary>Gets the ODF source cell range, when the source is a local cell range.</summary>
    public string? SourceRangeAddress => (string?)Element.Element(OdfNamespaces.Table + "source-cell-range")?
        .Attribute(OdfNamespaces.Table + "cell-range-address");

    /// <summary>Gets the ODF target cell range.</summary>
    public string? TargetRangeAddress => (string?)Element.Attribute(OdfNamespaces.Table + "target-range-address");

    /// <summary>Gets the source fields in document order.</summary>
    public IReadOnlyList<OdsDataPilotField> Fields => Element.Elements(OdfNamespaces.Table + "data-pilot-field")
        .Select(element => new OdsDataPilotField(element)).ToList();

    /// <summary>Whether imported grouping, field references, member selection, or nonlocal sources exceed the simple conversion subset.</summary>
    public bool HasAdvancedSettings => HasAdvancedSettingsIn(Element);

    internal static bool IsEditableElement(OdsDocument document, XElement element) {
        try {
            if (!ReferenceEquals(element.Parent?.Parent, document.SpreadsheetBody)
                || HasAdvancedSettingsIn(element)
                || string.IsNullOrWhiteSpace((string?)element.Attribute(OdfNamespaces.Table + "name"))) return false;
            SpreadsheetRangeReference source = ParseLocalRange((string?)element.Element(OdfNamespaces.Table + "source-cell-range")?
                .Attribute(OdfNamespaces.Table + "cell-range-address") ?? string.Empty, nameof(SourceRangeAddress));
            SpreadsheetRangeReference target = ParseLocalRange((string?)element.Attribute(OdfNamespaces.Table + "target-range-address")
                ?? string.Empty, nameof(TargetRangeAddress));
            OdsSheet? sourceSheet = document.GetSheet(source.Start.SheetName!);
            if (sourceSheet == null || document.GetSheet(target.Start.SheetName!) == null)
                return false;
            var pivot = new OdsDataPilotTable(document, element);
            IReadOnlyList<OdsDataPilotField> fields = pivot.Fields;
            if (fields.Count == 0) {
                long headerRow = source.Start.Row!.Value - 1;
                OdsRowRun? header = sourceSheet.RowRuns.FirstOrDefault(row =>
                    headerRow >= row.StartRow && headerRow - row.StartRow < row.RepeatCount);
                if (header == null) return false;
                _ = header.CellRuns;
            }
            return fields.All(field => pivot.SourceHeaderMatchesExactlyOnce(field.SourceFieldName));
        } catch (Exception exception) when (exception is ArgumentException or System.IO.InvalidDataException or OverflowException) {
            return false;
        }
    }

    private static bool HasAdvancedSettingsIn(XElement element) {
        XElement? source = element.Element(OdfNamespaces.Table + "source-cell-range");
        return source == null || source.HasElements
            || HasOtherAttributes(source, OdfNamespaces.Table + "cell-range-address")
            || HasOtherAttributes(element, OdfNamespaces.Table + "name",
                OdfNamespaces.Table + "target-range-address", OdfNamespaces.Table + "show-filter-button",
                OdfNamespaces.Table + "buttons")
            || element.Elements().Any(child => child.Name != OdfNamespaces.Table + "source-cell-range"
                && child.Name != OdfNamespaces.Table + "data-pilot-field")
            || element.Elements(OdfNamespaces.Table + "data-pilot-field").Any(IsAdvancedField);
    }

    private static bool IsAdvancedField(XElement element) {
        if (HasOtherAttributes(element, OdfNamespaces.Table + "source-field-name",
                OdfNamespaces.Table + "orientation", OdfNamespaces.Table + "function")) return true;
        if (!IsSupportedFieldShape((string?)element.Attribute(OdfNamespaces.Table + "source-field-name"),
                (string?)element.Attribute(OdfNamespaces.Table + "orientation"),
                (string?)element.Attribute(OdfNamespaces.Table + "function"))) return true;
        foreach (XElement level in element.Elements()) {
            if (level.Name != OdfNamespaces.Table + "data-pilot-level"
                || HasOtherAttributes(level, OdfNamespaces.Table + "show-empty")
                || (string?)level.Attribute(OdfNamespaces.Table + "show-empty") == "true") return true;
            foreach (XElement setting in level.Elements()) {
                if (setting.Name == OdfNamespaces.Table + "data-pilot-subtotals") {
                    if (HasOtherAttributes(setting) || setting.Elements().Count() != 1 || setting.Elements().Any(item =>
                        item.Name != OdfNamespaces.Table + "data-pilot-subtotal"
                        || (string?)item.Attribute(OdfNamespaces.Table + "function") != "auto"
                        || HasOtherAttributes(item, OdfNamespaces.Table + "function"))) return true;
                } else if (setting.Name == OdfNamespaces.Table + "data-pilot-sort-info") {
                    if (setting.HasElements || (string?)setting.Attribute(OdfNamespaces.Table + "sort-mode") != "manual"
                        || (string?)setting.Attribute(OdfNamespaces.Table + "order") != "ascending"
                        || HasOtherAttributes(setting, OdfNamespaces.Table + "sort-mode", OdfNamespaces.Table + "order")) return true;
                } else if (setting.Name == OdfNamespaces.Table + "data-pilot-layout-info") {
                    if (setting.HasElements || (string?)setting.Attribute(OdfNamespaces.Table + "layout-mode") != "outline-subtotals-top"
                        || (string?)setting.Attribute(OdfNamespaces.Table + "add-empty-lines") != "false"
                        || HasOtherAttributes(setting, OdfNamespaces.Table + "layout-mode", OdfNamespaces.Table + "add-empty-lines")) return true;
                } else return true;
            }
        }
        return false;
    }

    private static bool HasOtherAttributes(XElement element, params XName[] allowed) => element.Attributes()
        .Any(attribute => !attribute.IsNamespaceDeclaration && !allowed.Contains(attribute.Name));

    private static bool IsSupportedFieldShape(string? name, string? orientation, string? function) =>
        !string.IsNullOrWhiteSpace(name) &&
        (orientation switch {
            "data" => function is "sum" or "average" or "count" or "countnums" or "max" or "min"
                or "product" or "stdev" or "stdevp" or "var" or "varp",
            "row" or "column" or "page" => function is null or "auto",
            _ => false
        });

    /// <summary>Adds a source field to this table using an ODF orientation and optional aggregation function.</summary>
    public OdsDataPilotField AddField(string sourceFieldName, string orientation, string? function = null) {
        if (HasAdvancedSettings)
            throw new InvalidOperationException("Fields cannot be edited on a data pilot with advanced imported settings.");
        SpreadsheetRangeReference target;
        try {
            target = ParseLocalRange(TargetRangeAddress ?? string.Empty, nameof(TargetRangeAddress));
        } catch (ArgumentException) {
            throw new InvalidOperationException("Fields cannot be edited while the data pilot target range is invalid.");
        }
        if (!ReferenceEquals(Element.Parent?.Parent, _document.SpreadsheetBody)
            || _document.GetSheet(target.Start.SheetName!) == null)
            throw new InvalidOperationException("Fields cannot be edited while the data pilot target worksheet is missing.");
        if (Fields.Any(field => !SourceHeaderMatchesExactlyOnce(field.SourceFieldName)))
            throw new InvalidOperationException("Fields cannot be edited while an existing source header binding is missing.");
        if (string.IsNullOrWhiteSpace(sourceFieldName)) throw new ArgumentException("Source field name cannot be empty.", nameof(sourceFieldName));
        if (orientation != "row" && orientation != "column" && orientation != "data" && orientation != "page") {
            throw new ArgumentException("Orientation must be row, column, data, or page.", nameof(orientation));
        }
        if (orientation == "data") {
            if (function != "sum" && function != "average" && function != "count" && function != "countnums"
                && function != "max" && function != "min" && function != "product" && function != "stdev"
                && function != "stdevp" && function != "var" && function != "varp") {
                throw new ArgumentException("Data fields require a supported ODF aggregation function.", nameof(function));
            }
        } else if (function != null && function != "auto") {
            throw new ArgumentException("Non-data fields can use only the auto function.", nameof(function));
        }
        if (!SourceHeaderMatchesExactlyOnce(sourceFieldName)) {
            throw new ArgumentException("Source field name must match exactly one header in the source range.", nameof(sourceFieldName));
        }
        if (Fields.Any(field => string.Equals(field.SourceFieldName, sourceFieldName, StringComparison.Ordinal)
            && string.Equals(field.Orientation, orientation, StringComparison.Ordinal)
            && (orientation != "data" || string.Equals(field.Function, function, StringComparison.Ordinal)))) {
            throw new InvalidOperationException("The source field is already present in this orientation.");
        }
        var element = new XElement(OdfNamespaces.Table + "data-pilot-field",
            new XAttribute(OdfNamespaces.Table + "source-field-name", sourceFieldName),
            new XAttribute(OdfNamespaces.Table + "orientation", orientation));
        if (function != null) element.SetAttributeValue(OdfNamespaces.Table + "function", function);
        Element.Add(element);
        _document.MarkPartDirty("content.xml");
        return new OdsDataPilotField(element);
    }

    private bool SourceHeaderMatchesExactlyOnce(string fieldName) {
        SpreadsheetRangeReference range = ParseLocalRange(SourceRangeAddress ?? string.Empty, nameof(SourceRangeAddress));
        OdsSheet? sheet = _document.GetSheet(range.Start.SheetName!);
        if (sheet == null) return false;
        long headerRow = range.Start.Row!.Value - 1;
        long firstColumn = range.Start.Column!.Value - 1;
        long lastColumn = range.End!.Column!.Value - 1;
        bool matched = false;
        foreach (OdsRowRun row in sheet.RowRuns) {
            if (row.StartRow > headerRow) break;
            if (headerRow - row.StartRow >= row.RepeatCount) continue;
            foreach (OdsCellRun cell in row.CellRuns) {
                if (cell.StartColumn > lastColumn) break;
                long firstIncluded = Math.Max(firstColumn, cell.StartColumn);
                long cellEnd = cell.RepeatCount > long.MaxValue - cell.StartColumn
                    ? long.MaxValue : cell.StartColumn + cell.RepeatCount;
                long lastExclusive = Math.Min(lastColumn + 1, cellEnd);
                if (firstIncluded >= lastExclusive) continue;
                if (cell.Value.Kind != OdsCellValueKind.Empty
                    && string.Equals(cell.Text, fieldName, StringComparison.Ordinal)) {
                    if (matched || lastExclusive - firstIncluded > 1) return false;
                    matched = true;
                }
            }
            break;
        }
        return matched;
    }

    internal static SpreadsheetRangeReference ParseLocalRange(string address, string parameterName) {
        if (!SpreadsheetRangeReference.TryParse(address, SpreadsheetAddressDialect.OpenDocument,
                out SpreadsheetRangeReference? range) || !range!.IsRange || !range.Start.IsCell || !range.End!.IsCell
            || string.IsNullOrEmpty(range.Start.SheetName)
            || !string.Equals(range.Start.SheetName, range.End.SheetName, StringComparison.Ordinal)
            || range.End.Row < range.Start.Row || range.End.Column < range.Start.Column) {
            throw new ArgumentException("Address must be a rectangular range on one named worksheet.", parameterName);
        }
        return range;
    }
}

/// <summary>One field of an ODF data pilot table.</summary>
public sealed class OdsDataPilotField {
    internal XElement Element { get; }
    internal OdsDataPilotField(XElement element) { Element = element; }

    /// <summary>Gets the source header name.</summary>
    public string SourceFieldName => (string?)Element.Attribute(OdfNamespaces.Table + "source-field-name") ?? string.Empty;

    /// <summary>Gets the ODF field orientation.</summary>
    public string Orientation => (string?)Element.Attribute(OdfNamespaces.Table + "orientation") ?? string.Empty;

    /// <summary>Gets the ODF aggregation function, if present.</summary>
    public string? Function => (string?)Element.Attribute(OdfNamespaces.Table + "function");
}
