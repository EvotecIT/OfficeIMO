namespace OfficeIMO.OpenDocument;

/// <summary>Base class for an XML-backed ODP drawing shape.</summary>
public abstract class OdpShape : OdfShape {
    internal OdpShape(OdpPresentation presentation, XElement element) : base(presentation, element) { Presentation = presentation; }
    internal OdpPresentation Presentation { get; }
    /// <summary>Whether the shape is hidden from the normal presentation view.</summary>
    public bool Hidden {
        get => (string?)Element.Attribute(OdfNamespaces.Presentation + "visibility") == "hidden";
        set { Element.SetAttributeValue(OdfNamespaces.Presentation + "visibility", value ? "hidden" : null); Dirty(); }
    }
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
