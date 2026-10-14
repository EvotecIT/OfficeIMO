namespace OfficeIMO.OpenDocument;

/// <summary>Basic native ODF fields with cached display text.</summary>
public enum OdfTextFieldKind {
    /// <summary>Current page number.</summary>
    PageNumber,
    /// <summary>Total page count.</summary>
    PageCount,
    /// <summary>Date.</summary>
    Date,
    /// <summary>Time.</summary>
    Time
}

/// <summary>A basic native field. Other field kinds remain preserved as opaque inline XML.</summary>
public sealed partial class OdfTextField {
    private readonly OdfDocument _document;
    private readonly XElement _element;
    internal OdfTextField(OdfDocument document, XElement element) { _document = document; _element = element; }
    /// <summary>Native field kind.</summary>
    public OdfTextFieldKind Kind => TryGetKind(_element.Name, out OdfTextFieldKind kind) ? kind : throw new InvalidOperationException("Not a basic ODF field.");
    /// <summary>Cached display text, which a native application may refresh. Assignment preserves field attributes.</summary>
    public string DisplayText { get => OdfTextCodec.Read(_element); set { _element.Value = value ?? string.Empty; _document.MarkPartDirty(_document.GetPartPath(_element)); } }
    /// <summary>Whether the cached value is fixed. Page count is always dynamic.</summary>
    public bool IsFixed {
        get => OdfBoolean.TryParseXml((string?)_element.Attribute(OdfNamespaces.Text + "fixed"), out bool value) && value;
        set {
            if (Kind == OdfTextFieldKind.PageCount && value) throw new NotSupportedException("Page-count fields cannot be fixed.");
            _element.SetAttributeValue(OdfNamespaces.Text + "fixed", value ? "true" : null); _document.MarkPartDirty(_document.GetPartPath(_element));
        }
    }
    /// <summary>Returns a detached copy including native date/time values, style names and offsets.</summary>
    public XElement ToXml() => new XElement(_element);

    internal static bool IsField(XName name) => TryGetKind(name, out _);
    internal static bool TryGetKind(XName name, out OdfTextFieldKind kind) {
        if (name == OdfNamespaces.Text + "page-number") kind = OdfTextFieldKind.PageNumber;
        else if (name == OdfNamespaces.Text + "page-count") kind = OdfTextFieldKind.PageCount;
        else if (name == OdfNamespaces.Text + "date") kind = OdfTextFieldKind.Date;
        else if (name == OdfNamespaces.Text + "time") kind = OdfTextFieldKind.Time;
        else { kind = default; return false; }
        return true;
    }
    internal static XElement Create(OdfTextFieldKind kind, string? text) {
        string token = kind switch {
            OdfTextFieldKind.PageNumber => "page-number", OdfTextFieldKind.PageCount => "page-count",
            OdfTextFieldKind.Date => "date", OdfTextFieldKind.Time => "time", _ => throw new ArgumentOutOfRangeException(nameof(kind))
        };
        return new XElement(OdfNamespaces.Text + token, text ?? string.Empty);
    }
}
