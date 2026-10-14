namespace OfficeIMO.OpenDocument;

/// <summary>Document fields supported by the native ODT text surface.</summary>
public enum OdtFieldKind {
    /// <summary>The current page number.</summary>
    PageNumber,
    /// <summary>The document page count.</summary>
    PageCount,
    /// <summary>The current date.</summary>
    Date,
    /// <summary>The current time.</summary>
    Time
}

/// <summary>An XML-backed ODT field with cached display text.</summary>
public sealed class OdtField {
    private readonly OdtDocument _document;
    private XElement _element;
    private Func<XElement>? _materializeForEdit;
    private readonly string _partPath;

    internal OdtField(OdtDocument document, XElement element, string partPath, Func<XElement>? materializeForEdit = null) {
        _document = document;
        _element = element;
        _partPath = partPath;
        _materializeForEdit = materializeForEdit;
    }

    /// <summary>The field's native ODF kind.</summary>
    public OdtFieldKind Kind => _element.Name.LocalName switch {
        "page-number" => OdtFieldKind.PageNumber,
        "page-count" => OdtFieldKind.PageCount,
        "date" => OdtFieldKind.Date,
        "time" => OdtFieldKind.Time,
        _ => throw new InvalidOperationException("The element is not a supported ODT field.")
    };

    /// <summary>Cached text shown until the document application refreshes the field.</summary>
    public string DisplayText {
        get => OdfTextCodec.Read(_element);
        set {
            EnsureMaterialized();
            // ODF fields have text-only content in the schema. Whitespace elements
            // used by paragraphs would turn a simple field into invalid XML.
            _element.Value = value ?? string.Empty;
            _document.MarkPartDirty(_partPath);
        }
    }

    /// <summary>Whether the displayed value is fixed instead of refreshed. Page count is always dynamic.</summary>
    public bool IsFixed {
        get => OdfBoolean.TryParseXml((string?)_element.Attribute(OdfNamespaces.Text + "fixed"),
            out bool value) && value;
        set {
            EnsureMaterialized();
            if (Kind == OdtFieldKind.PageCount && value)
                throw new NotSupportedException("ODT page-count fields cannot be fixed.");
            _element.SetAttributeValue(OdfNamespaces.Text + "fixed", value ? "true" : null);
            _document.MarkPartDirty(_partPath);
        }
    }

    /// <summary>True when the field has no style, offset, value override, or unrecognized property.</summary>
    public bool IsBasic => IsBasicElement(_element);

    internal static bool IsBasicElement(XElement element) =>
        TryGetKind(element.Name, out OdtFieldKind kind) && !element.HasElements &&
        element.Attributes().All(attribute => attribute.IsNamespaceDeclaration ||
            attribute.Name == OdfNamespaces.Text + "fixed" &&
            kind != OdtFieldKind.PageCount &&
            OdfBoolean.TryParseXml(attribute.Value, out _));

    internal XElement Element => _element;

    internal static bool TryGetKind(XName name, out OdtFieldKind kind) {
        bool found = OdfTextField.TryGetKind(name, out OdfTextFieldKind nativeKind);
        kind = (OdtFieldKind)nativeKind; return found;
    }

    internal static XElement CreateElement(OdtFieldKind kind, string? displayText) {
        return OdfTextField.Create((OdfTextFieldKind)kind, displayText);
    }
    private XElement EnsureMaterialized() {
        if (_materializeForEdit != null) {
            _element = _materializeForEdit();
            _materializeForEdit = null;
        }
        return _element;
    }

}
