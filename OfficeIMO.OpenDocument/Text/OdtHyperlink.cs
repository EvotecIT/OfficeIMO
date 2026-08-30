namespace OfficeIMO.OpenDocument;

/// <summary>An XML-backed ODT hyperlink. Targets are preserved and never fetched.</summary>
public sealed class OdtHyperlink {
    private readonly OdtDocument _document;
    private readonly XElement _element;
    private readonly string _partPath;

    internal OdtHyperlink(OdtDocument document, XElement element, string partPath = "content.xml") {
        _document = document;
        _element = element;
        _partPath = partPath;
    }

    /// <summary>Decoded display text.</summary>
    public string Text {
        get => OdfTextCodec.Read(_element);
        set { OdfTextCodec.Replace(_element, value); Dirty(); }
    }
    /// <summary>Changes the hyperlink display text casing while preserving its target and text style.</summary>
    public OdtHyperlink TransformTextCase(OfficeIMO.Drawing.OfficeTextCase textCase, System.Globalization.CultureInfo? culture = null) {
        OdfTextCodec.TransformTextCase(_element, textCase, culture);
        Dirty();
        return this;
    }
    /// <summary>Link target.</summary>
    public string Href {
        get => (string?)_element.Attribute(OdfNamespaces.XLink + "href") ?? string.Empty;
        set {
            if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("Hyperlink target cannot be empty.", nameof(value));
            _element.SetAttributeValue(OdfNamespaces.XLink + "href", value);
            Dirty();
        }
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
    /// <summary>Explicit or inherited underline state.</summary>
    public bool? Underline { get => Resolve(style => style.Underline); set => EnsureStyle().Underline = value; }
    /// <summary>Explicit or inherited native underline style.</summary>
    public OdfTextDecorationStyle? UnderlineStyle { get => Resolve(style => style.UnderlineStyle); set => EnsureStyle().UnderlineStyle = value; }
    /// <summary>Explicit or inherited native underline line count.</summary>
    public OdfTextDecorationType? UnderlineType { get => Resolve(style => style.UnderlineType); set => EnsureStyle().UnderlineType = value; }
    /// <summary>Whether the effective underline uses a non-solid ODF decoration style.</summary>
    public bool UsesNonSolidUnderlineStyle => Resolve(style => style.UsesNonSolidUnderlineStyle) == true;
    /// <summary>Explicit or inherited strike-through state.</summary>
    public bool? StrikeThrough { get => Resolve(style => style.StrikeThrough); set => EnsureStyle().StrikeThrough = value; }
    /// <summary>Explicit or inherited native line-through style.</summary>
    public OdfTextDecorationStyle? LineThroughStyle { get => Resolve(style => style.LineThroughStyle); set => EnsureStyle().LineThroughStyle = value; }
    /// <summary>Explicit or inherited native line-through line count.</summary>
    public OdfTextDecorationType? LineThroughType { get => Resolve(style => style.LineThroughType); set => EnsureStyle().LineThroughType = value; }
    /// <summary>Whether the effective line-through uses a non-solid ODF decoration style.</summary>
    public bool UsesNonSolidLineThroughStyle => Resolve(style => style.UsesNonSolidLineThroughStyle) == true;
    /// <summary>Explicit or inherited font size.</summary>
    public OdfLength? FontSize { get => Resolve(style => style.FontSize); set => EnsureStyle().FontSize = value; }
    /// <summary>Explicit or inherited baseline placement.</summary>
    public OdfTextPosition? TextPosition { get => Resolve(style => style.TextPosition); set => EnsureStyle().TextPosition = value; }
    /// <summary>Explicit or inherited display-time case transformation.</summary>
    public OdfTextTransform? TextTransform { get => Resolve(style => style.TextTransform); set => EnsureStyle().TextTransform = value; }
    /// <summary>Explicit or inherited small-cap display formatting.</summary>
    public bool? SmallCaps { get => Resolve(style => style.SmallCaps); set => EnsureStyle().SmallCaps = value; }
    /// <summary>Explicit or inherited font family.</summary>
    public string? FontFamily { get => ResolveReference(style => style.FontFamily); set => EnsureStyle().FontFamily = value; }
    /// <summary>Explicit or inherited text color.</summary>
    public OdfColor? Color { get => Resolve(style => style.Color); set => EnsureStyle().Color = value; }
    /// <summary>Explicit or inherited text background color.</summary>
    public OdfColor? BackgroundColor {
        get {
            OdfStyle? style = StyleName == null ? null : _document.Styles.FindInPart(
                OdfStyleFamily.Text, StyleName, _partPath);
            return _document.Styles.ResolveTextBackgroundColor(style);
        }
        set => EnsureStyle().TextBackgroundColor = value;
    }

    private OdfStyle EnsureStyle() => _document.Styles.EnsureAutomaticStyle(
        _element, OdfNamespaces.Text + "style-name", OdfStyleFamily.Text, "ofL", _partPath);

    private T? Resolve<T>(Func<OdfStyle, T?> selector) where T : struct {
        OdfStyle? style = StyleName == null ? null : _document.Styles.FindInPart(OdfStyleFamily.Text, StyleName, _partPath);
        if (style == null) return null;
        foreach (OdfStyle candidate in _document.Styles.Resolve(style)) {
            T? value = selector(candidate);
            if (value.HasValue) return value;
        }
        return null;
    }

    private string? ResolveReference(Func<OdfStyle, string?> selector) {
        OdfStyle? style = StyleName == null ? null : _document.Styles.FindInPart(OdfStyleFamily.Text, StyleName, _partPath);
        if (style == null) return null;
        foreach (OdfStyle candidate in _document.Styles.Resolve(style)) {
            string? value = selector(candidate);
            if (value != null) return value;
        }
        return null;
    }

    private void Dirty() => _document.MarkPartDirty(_partPath);
}
