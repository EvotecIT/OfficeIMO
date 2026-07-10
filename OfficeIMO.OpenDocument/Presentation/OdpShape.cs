namespace OfficeIMO.OpenDocument;

/// <summary>Base class for an XML-backed ODP drawing shape.</summary>
public abstract class OdpShape {
    internal OdpShape(OdpPresentation presentation, XElement element) { Presentation = presentation; Element = element; }
    /// <summary>Shape name.</summary>
    public string Name {
        get => (string?)Element.Attribute(OdfNamespaces.Draw + "name") ?? string.Empty;
        set { Element.SetAttributeValue(OdfNamespaces.Draw + "name", value); Dirty(); }
    }
    /// <summary>Whether the shape is hidden from the normal presentation view.</summary>
    public bool Hidden {
        get => (string?)Element.Attribute(OdfNamespaces.Presentation + "visibility") == "hidden";
        set { Element.SetAttributeValue(OdfNamespaces.Presentation + "visibility", value ? "hidden" : null); Dirty(); }
    }
    /// <summary>Raw ODF/SVG transform expression.</summary>
    public string? Transform {
        get => (string?)Element.Attribute(OdfNamespaces.Draw + "transform");
        set { Element.SetAttributeValue(OdfNamespaces.Draw + "transform", value); Dirty(); }
    }
    /// <summary>Solid shape fill color.</summary>
    public OdfColor? FillColor {
        get => ReadGraphicColor(OdfNamespaces.Draw + "fill-color");
        set {
            OdfStyle style = EnsureGraphicStyle();
            style.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "fill", value.HasValue ? "solid" : "none");
            style.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "fill-color", value?.ToString());
        }
    }
    /// <summary>Solid shape stroke color.</summary>
    public OdfColor? StrokeColor {
        get => ReadGraphicColor(OdfNamespaces.Svg + "stroke-color");
        set {
            OdfStyle style = EnsureGraphicStyle();
            style.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "stroke", value.HasValue ? "solid" : "none");
            style.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Svg + "stroke-color", value?.ToString());
        }
    }
    /// <summary>Shape stroke width.</summary>
    public OdfLength? StrokeWidth {
        get {
            string? value = (string?)GetGraphicStyle()?.Element.Element(OdfNamespaces.Style + "graphic-properties")?.Attribute(OdfNamespaces.Svg + "stroke-width");
            return value == null ? (OdfLength?)null : OdfLength.Parse(value);
        }
        set => EnsureGraphicStyle().SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Svg + "stroke-width", value?.ToString());
    }
    /// <summary>Position and size for shapes exposing SVG bounds.</summary>
    public virtual OdfRect Bounds {
        get => new OdfRect(ReadLength("x"), ReadLength("y"), ReadLength("width"), ReadLength("height"));
        set { ApplyBounds(Element, value); Dirty(); }
    }
    internal OdpPresentation Presentation { get; }
    internal XElement Element { get; }
    internal static OdpShape? Wrap(OdpPresentation presentation, XElement element) {
        if (element.Name == OdfNamespaces.Draw + "frame") {
            if (element.Element(OdfNamespaces.Draw + "text-box") != null) return new OdpTextBox(presentation, element);
            if (element.Element(OdfNamespaces.Draw + "image") != null) return new OdpImage(presentation, element);
            if (element.Element(OdfNamespaces.Table + "table") != null) return new OdpTable(presentation, element);
        }
        if (element.Name == OdfNamespaces.Draw + "rect") return new OdpRectangle(presentation, element);
        if (element.Name == OdfNamespaces.Draw + "ellipse") return new OdpEllipse(presentation, element);
        if (element.Name == OdfNamespaces.Draw + "line") return new OdpLine(presentation, element);
        if (element.Name == OdfNamespaces.Draw + "g") return new OdpGroup(presentation, element);
        return null;
    }
    internal static void ApplyBounds(XElement element, OdfRect bounds) {
        element.SetAttributeValue(OdfNamespaces.Svg + "x", bounds.X.ToString());
        element.SetAttributeValue(OdfNamespaces.Svg + "y", bounds.Y.ToString());
        element.SetAttributeValue(OdfNamespaces.Svg + "width", bounds.Width.ToString());
        element.SetAttributeValue(OdfNamespaces.Svg + "height", bounds.Height.ToString());
    }
    internal void Dirty() => Presentation.MarkPartDirty("content.xml");
    internal OdfStyle EnsureGraphicStyle() => Presentation.Styles.EnsureAutomaticStyle(Element, OdfNamespaces.Draw + "style-name", OdfStyleFamily.Graphic, "ofGr");
    private OdfStyle? GetGraphicStyle() {
        string? name = (string?)Element.Attribute(OdfNamespaces.Draw + "style-name");
        return name == null ? null : Presentation.Styles.Find(OdfStyleFamily.Graphic, name);
    }
    private OdfColor? ReadGraphicColor(XName name) {
        string? value = (string?)GetGraphicStyle()?.Element.Element(OdfNamespaces.Style + "graphic-properties")?.Attribute(name);
        return value == null ? (OdfColor?)null : OdfColor.Parse(value);
    }
    private OdfLength ReadLength(string localName) => OdfLength.Parse((string?)Element.Attribute(OdfNamespaces.Svg + localName) ?? "0cm");
}

/// <summary>An ODP rectangle.</summary>
public sealed class OdpRectangle : OdpShape {
    internal OdpRectangle(OdpPresentation presentation, XElement element) : base(presentation, element) { }
    internal static OdpRectangle Create(OdpPresentation presentation, OdfRect bounds, string name) {
        var element = new XElement(OdfNamespaces.Draw + "rect", new XAttribute(OdfNamespaces.Draw + "name", name)); ApplyBounds(element, bounds); return new OdpRectangle(presentation, element);
    }
}

/// <summary>An ODP ellipse.</summary>
public sealed class OdpEllipse : OdpShape {
    internal OdpEllipse(OdpPresentation presentation, XElement element) : base(presentation, element) { }
    internal static OdpEllipse Create(OdpPresentation presentation, OdfRect bounds, string name) {
        var element = new XElement(OdfNamespaces.Draw + "ellipse", new XAttribute(OdfNamespaces.Draw + "name", name)); ApplyBounds(element, bounds); return new OdpEllipse(presentation, element);
    }
}

/// <summary>An ODP line.</summary>
public sealed class OdpLine : OdpShape {
    internal OdpLine(OdpPresentation presentation, XElement element) : base(presentation, element) { }
    /// <inheritdoc />
    public override OdfRect Bounds { get => base.Bounds; set => throw new NotSupportedException("Set line endpoints instead of rectangular bounds."); }
    /// <summary>First horizontal endpoint.</summary>
    public OdfLength X1 { get => ReadEndpoint("x1"); set => SetEndpoint("x1", value); }
    /// <summary>First vertical endpoint.</summary>
    public OdfLength Y1 { get => ReadEndpoint("y1"); set => SetEndpoint("y1", value); }
    /// <summary>Second horizontal endpoint.</summary>
    public OdfLength X2 { get => ReadEndpoint("x2"); set => SetEndpoint("x2", value); }
    /// <summary>Second vertical endpoint.</summary>
    public OdfLength Y2 { get => ReadEndpoint("y2"); set => SetEndpoint("y2", value); }
    internal static OdpLine Create(OdpPresentation presentation, OdfLength x1, OdfLength y1, OdfLength x2, OdfLength y2, string name) {
        return new OdpLine(presentation, new XElement(OdfNamespaces.Draw + "line",
            new XAttribute(OdfNamespaces.Draw + "name", name), new XAttribute(OdfNamespaces.Svg + "x1", x1.ToString()),
            new XAttribute(OdfNamespaces.Svg + "y1", y1.ToString()), new XAttribute(OdfNamespaces.Svg + "x2", x2.ToString()),
            new XAttribute(OdfNamespaces.Svg + "y2", y2.ToString())));
    }
    private OdfLength ReadEndpoint(string name) => OdfLength.Parse((string?)Element.Attribute(OdfNamespaces.Svg + name) ?? "0cm");
    private void SetEndpoint(string name, OdfLength value) { Element.SetAttributeValue(OdfNamespaces.Svg + name, value.ToString()); Dirty(); }
}
