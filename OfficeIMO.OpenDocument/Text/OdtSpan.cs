namespace OfficeIMO.OpenDocument;

/// <summary>An XML-backed ODT inline span.</summary>
public sealed class OdtSpan {
    private readonly OdtDocument _document;
    private readonly XElement _element;
    private readonly string _partPath;

    internal OdtSpan(OdtDocument document, XElement element, string partPath = "content.xml") {
        _document = document;
        _element = element;
        _partPath = partPath;
    }

    /// <summary>Decoded span text.</summary>
    public string Text {
        get => OdfTextCodec.Read(_element);
        set { OdfTextCodec.Replace(_element, value); Dirty(); }
    }
    /// <summary>Referenced text style name.</summary>
    public string? StyleName {
        get => (string?)_element.Attribute(OdfNamespaces.Text + "style-name");
        set { _element.SetAttributeValue(OdfNamespaces.Text + "style-name", value); Dirty(); }
    }
    /// <summary>Explicit or inherited bold state.</summary>
    public bool? Bold { get => Resolve(style => style.Bold); set => EnsureStyle().Bold = value; }
    /// <summary>Explicit or inherited italic state.</summary>
    public bool? Italic { get => Resolve(style => style.Italic); set => EnsureStyle().Italic = value; }
    /// <summary>Explicit or inherited font size.</summary>
    public OdfLength? FontSize { get => Resolve(style => style.FontSize); set => EnsureStyle().FontSize = value; }
    /// <summary>Explicit or inherited text color.</summary>
    public OdfColor? Color { get => Resolve(style => style.Color); set => EnsureStyle().Color = value; }

    /// <summary>Appends decoded plain text.</summary>
    public OdtSpan AddText(string text) { OdfTextCodec.Append(_element, text); Dirty(); return this; }

    private OdfStyle EnsureStyle() => _document.Styles.EnsureAutomaticStyle(
        _element, OdfNamespaces.Text + "style-name", OdfStyleFamily.Text, "ofT", _partPath);

    private T? Resolve<T>(Func<OdfStyle, T?> selector) where T : struct {
        OdfStyle? style = StyleName == null ? null : _document.Styles.Find(OdfStyleFamily.Text, StyleName);
        if (style == null) return null;
        foreach (OdfStyle candidate in _document.Styles.Resolve(style)) {
            T? value = selector(candidate);
            if (value.HasValue) return value;
        }
        return null;
    }

    private void Dirty() => _document.MarkPartDirty(_partPath);
}
