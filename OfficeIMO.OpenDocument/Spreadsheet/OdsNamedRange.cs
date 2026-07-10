namespace OfficeIMO.OpenDocument;

/// <summary>An XML-backed workbook named range.</summary>
public sealed class OdsNamedRange {
    private readonly OdsDocument _document;
    private readonly XElement _element;

    internal OdsNamedRange(OdsDocument document, XElement element) { _document = document; _element = element; }

    /// <summary>Range name.</summary>
    public string Name => (string?)_element.Attribute(OdfNamespaces.Table + "name") ?? string.Empty;
    /// <summary>ODF cell range address.</summary>
    public string CellRangeAddress {
        get => (string?)_element.Attribute(OdfNamespaces.Table + "cell-range-address") ?? string.Empty;
        set { _element.SetAttributeValue(OdfNamespaces.Table + "cell-range-address", value); Dirty(); }
    }
    /// <summary>ODF base cell address.</summary>
    public string BaseCellAddress {
        get => (string?)_element.Attribute(OdfNamespaces.Table + "base-cell-address") ?? string.Empty;
        set { _element.SetAttributeValue(OdfNamespaces.Table + "base-cell-address", value); Dirty(); }
    }

    private void Dirty() => _document.MarkPartDirty("content.xml");
}

/// <summary>An XML-backed spreadsheet content validation rule.</summary>
public sealed class OdsValidation {
    private readonly OdsDocument _document;
    private readonly XElement _element;

    internal OdsValidation(OdsDocument document, XElement element) { _document = document; _element = element; }

    /// <summary>Validation name.</summary>
    public string Name => (string?)_element.Attribute(OdfNamespaces.Table + "name") ?? string.Empty;
    /// <summary>Preserved ODF validation condition.</summary>
    public string Condition {
        get => (string?)_element.Attribute(OdfNamespaces.Table + "condition") ?? string.Empty;
        set { _element.SetAttributeValue(OdfNamespaces.Table + "condition", value); Dirty(); }
    }
    /// <summary>Whether empty cells satisfy the rule.</summary>
    public bool AllowEmptyCell {
        get => (string?)_element.Attribute(OdfNamespaces.Table + "allow-empty-cell") != "false";
        set { _element.SetAttributeValue(OdfNamespaces.Table + "allow-empty-cell", value ? "true" : "false"); Dirty(); }
    }

    private void Dirty() => _document.MarkPartDirty("content.xml");
}
