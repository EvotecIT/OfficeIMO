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
    public bool HasAdvancedSettings {
        get {
            XElement? source = Element.Element(OdfNamespaces.Table + "source-cell-range");
            return source == null || source.HasElements
                || HasOtherAttributes(source, OdfNamespaces.Table + "cell-range-address")
                || HasOtherAttributes(Element, OdfNamespaces.Table + "name",
                    OdfNamespaces.Table + "target-range-address", OdfNamespaces.Table + "show-filter-button",
                    OdfNamespaces.Table + "buttons")
                || Element.Elements().Any(child => child.Name != OdfNamespaces.Table + "source-cell-range"
                    && child.Name != OdfNamespaces.Table + "data-pilot-field")
                || Element.Elements(OdfNamespaces.Table + "data-pilot-field").Any(IsAdvancedField);
        }
    }

    private static bool IsAdvancedField(XElement element) {
        if (HasOtherAttributes(element, OdfNamespaces.Table + "source-field-name",
                OdfNamespaces.Table + "orientation", OdfNamespaces.Table + "function")) return true;
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

    /// <summary>Adds a source field to this table using an ODF orientation and optional aggregation function.</summary>
    public OdsDataPilotField AddField(string sourceFieldName, string orientation, string? function = null) {
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
        if (Fields.Any(field => string.Equals(field.SourceFieldName, sourceFieldName, StringComparison.Ordinal)
            && string.Equals(field.Orientation, orientation, StringComparison.Ordinal))) {
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
