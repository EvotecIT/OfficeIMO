namespace OfficeIMO.OpenDocument;

/// <summary>An XML-backed ODP master page.</summary>
public sealed class OdpMasterPage {
    private readonly OdpPresentation _presentation;
    private readonly XElement _element;
    internal OdpMasterPage(OdpPresentation presentation, XElement element) { _presentation = presentation; _element = element; }
    /// <summary>Master page name.</summary>
    public string Name => (string?)_element.Attribute(OdfNamespaces.Style + "name") ?? string.Empty;
    /// <summary>Page layout name.</summary>
    public string? PageLayoutName => (string?)_element.Attribute(OdfNamespaces.Style + "page-layout-name");
    /// <summary>Background fill color.</summary>
    public OdfColor? BackgroundColor {
        get {
            string? styleName = (string?)_element.Attribute(OdfNamespaces.Draw + "style-name");
            OdfStyle? style = styleName == null ? null : _presentation.Styles.FindInPart(OdfStyleFamily.DrawingPage, styleName, "styles.xml");
            string? value = (string?)style?.Element.Element(OdfNamespaces.Style + "drawing-page-properties")?.Attribute(OdfNamespaces.Draw + "fill-color");
            return value == null ? (OdfColor?)null : OdfColor.Parse(value);
        }
        set {
            OdfStyle style = _presentation.Styles.EnsureAutomaticStyle(_element, OdfNamespaces.Draw + "style-name", OdfStyleFamily.DrawingPage, "ofMaster", "styles.xml");
            style.SetProperty(OdfNamespaces.Style + "drawing-page-properties", OdfNamespaces.Draw + "fill", value.HasValue ? "solid" : null);
            style.SetProperty(OdfNamespaces.Style + "drawing-page-properties", OdfNamespaces.Draw + "fill-color", value?.ToString());
        }
    }
}

/// <summary>An XML-backed ODP presentation page layout.</summary>
public sealed class OdpPresentationLayout {
    private readonly OdpPresentation _presentation;
    private readonly XElement _element;
    internal OdpPresentationLayout(OdpPresentation presentation, XElement element) { _presentation = presentation; _element = element; }
    /// <summary>Layout name.</summary>
    public string Name => (string?)_element.Attribute(OdfNamespaces.Style + "name") ?? string.Empty;
    /// <summary>Adds a presentation placeholder using an ODF object token such as <c>title</c> or <c>outline</c>.</summary>
    public void AddPlaceholder(string objectType, OdfRect bounds) {
        if (string.IsNullOrWhiteSpace(objectType)) throw new ArgumentException("Placeholder type cannot be empty.", nameof(objectType));
        _element.Add(new XElement(OdfNamespaces.Presentation + "placeholder",
            new XAttribute(OdfNamespaces.Presentation + "object", objectType),
            new XAttribute(OdfNamespaces.Svg + "x", bounds.X.ToString()),
            new XAttribute(OdfNamespaces.Svg + "y", bounds.Y.ToString()),
            new XAttribute(OdfNamespaces.Svg + "width", bounds.Width.ToString()),
            new XAttribute(OdfNamespaces.Svg + "height", bounds.Height.ToString())));
        _presentation.MarkPartDirty("styles.xml");
    }
}

internal sealed class OdpPageLayout {
    internal OdpPageLayout(OdpPresentation presentation, XElement element) {
        Element = element; Properties = new OdpPageLayoutProperties(presentation, element);
    }
    internal XElement Element { get; }
    internal string Name => (string?)Element.Attribute(OdfNamespaces.Style + "name") ?? string.Empty;
    internal OdpPageLayoutProperties Properties { get; }
}

internal sealed class OdpPageLayoutProperties {
    private readonly OdpPresentation _presentation;
    private readonly XElement _element;
    internal OdpPageLayoutProperties(OdpPresentation presentation, XElement pageLayout) {
        _presentation = presentation;
        _element = pageLayout.Element(OdfNamespaces.Style + "page-layout-properties") ?? new XElement(OdfNamespaces.Style + "page-layout-properties");
        if (_element.Parent == null) { pageLayout.Add(_element); presentation.MarkPartDirty("styles.xml"); }
    }
    internal OdfLength ReadLength(XName name, string fallback) => OdfLength.Parse((string?)_element.Attribute(name) ?? fallback);
    internal void SetLength(XName name, OdfLength value) { _element.SetAttributeValue(name, value.ToString()); _presentation.MarkPartDirty("styles.xml"); }
}
