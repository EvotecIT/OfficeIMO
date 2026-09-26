using OfficeIMO.Spreadsheet;

namespace OfficeIMO.OpenDocument;

/// <summary>An XML-backed conditional mapping from one ODF style to another.</summary>
public sealed class OdfStyleMap {
    private readonly OdfDocument _document;
    private readonly XElement _element;
    private readonly string _partPath;

    internal OdfStyleMap(OdfDocument document, XElement element, string partPath) {
        _document = document;
        _element = element;
        _partPath = partPath;
    }

    /// <summary>Preserved ODF condition expression.</summary>
    public string Condition {
        get => (string?)_element.Attribute(OdfNamespaces.Style + "condition") ?? string.Empty;
        set {
            if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("A condition is required.", nameof(value));
            SetAttribute(OdfNamespaces.Style + "condition", value);
        }
    }

    /// <summary>Name of the style applied when the condition is true.</summary>
    public string ApplyStyleName {
        get => (string?)_element.Attribute(OdfNamespaces.Style + "apply-style-name") ?? string.Empty;
        set {
            if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("An applied style name is required.", nameof(value));
            SetAttribute(OdfNamespaces.Style + "apply-style-name", value);
        }
    }

    /// <summary>Optional spreadsheet base cell for relative references in the condition.</summary>
    public string? BaseCellAddress {
        get => (string?)_element.Attribute(OdfNamespaces.Style + "base-cell-address");
        set {
            ValidateBaseCellAddress(value);
            SetAttribute(OdfNamespaces.Style + "base-cell-address", value);
        }
    }

    /// <summary>Removes this mapping from its owner style.</summary>
    public bool Remove() {
        if (_element.Parent == null) return false;
        _element.Remove();
        _document.MarkPartDirty(_partPath);
        return true;
    }

    internal static void ValidateBaseCellAddress(string? value) {
        if (!IsValidBaseCellAddress(value)) {
            throw new ArgumentException("Base cell address must identify one sheet-qualified OpenDocument cell.", nameof(value));
        }
    }

    internal static bool IsValidBaseCellAddress(string? value) {
        if (value == null) return true;
        return SpreadsheetRangeReference.TryParse(value, SpreadsheetAddressDialect.OpenDocument,
                out SpreadsheetRangeReference? reference)
            && reference!.End == null && reference.Start.IsCell && reference.Start.SheetName != null;
    }

    private void SetAttribute(XName name, string? value) {
        _element.SetAttributeValue(name, value);
        _document.MarkPartDirty(_partPath);
    }
}
