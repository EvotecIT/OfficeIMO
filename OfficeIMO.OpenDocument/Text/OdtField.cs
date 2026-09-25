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
    private readonly XElement _element;
    private readonly string _partPath;

    internal OdtField(OdtDocument document, XElement element, string partPath) {
        _document = document;
        _element = element;
        _partPath = partPath;
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
            // ODF fields have text-only content in the schema. Whitespace elements
            // used by paragraphs would turn a simple field into invalid XML.
            _element.Value = value ?? string.Empty;
            _document.MarkPartDirty(_partPath);
        }
    }

    /// <summary>Whether the displayed value is fixed instead of refreshed. Page count is always dynamic.</summary>
    public bool IsFixed {
        get => string.Equals((string?)_element.Attribute(OdfNamespaces.Text + "fixed"), "true",
            StringComparison.OrdinalIgnoreCase);
        set {
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
        element.Attributes().All(attribute => attribute.Name == OdfNamespaces.Text + "fixed" &&
            kind != OdtFieldKind.PageCount &&
            (attribute.Value == "true" || attribute.Value == "false"));

    internal XElement Element => _element;

    internal static bool TryGetKind(XName name, out OdtFieldKind kind) {
        if (name == OdfNamespaces.Text + "page-number") kind = OdtFieldKind.PageNumber;
        else if (name == OdfNamespaces.Text + "page-count") kind = OdtFieldKind.PageCount;
        else if (name == OdfNamespaces.Text + "date") kind = OdtFieldKind.Date;
        else if (name == OdfNamespaces.Text + "time") kind = OdtFieldKind.Time;
        else { kind = default; return false; }
        return true;
    }

    internal static XElement CreateElement(OdtFieldKind kind, string? displayText) {
        XName name = kind switch {
            OdtFieldKind.PageNumber => OdfNamespaces.Text + "page-number",
            OdtFieldKind.PageCount => OdfNamespaces.Text + "page-count",
            OdtFieldKind.Date => OdfNamespaces.Text + "date",
            OdtFieldKind.Time => OdfNamespaces.Text + "time",
            _ => throw new ArgumentOutOfRangeException(nameof(kind))
        };
        var element = new XElement(name);
        element.Value = displayText ?? string.Empty;
        return element;
    }
}
