namespace OfficeIMO.OpenDocument;

/// <summary>An XML-backed ODP shape group.</summary>
public sealed class OdpGroup : OdpShape {
    internal OdpGroup(OdpPresentation presentation, XElement element) : base(presentation, element) { }
    /// <inheritdoc />
    public override OdfRect Bounds { get => base.Bounds; set => throw new NotSupportedException("Groups use child geometry and draw:transform."); }
    /// <summary>Child shapes.</summary>
    public IReadOnlyList<OdpShape> Shapes => Element.Elements().Select(element => Wrap(Presentation, element))
        .Where(shape => shape != null).Select(shape => shape!).ToList();
    /// <summary>Adds a rectangle to the group.</summary>
    public OdpRectangle AddRectangle(OdfRect bounds, string? name = null) {
        OdpRectangle shape = OdpRectangle.Create(Presentation, bounds, name ?? NextName("Rectangle")); Element.Add(shape.Element); Dirty(); return shape;
    }
    /// <summary>Adds an ellipse to the group.</summary>
    public OdpEllipse AddEllipse(OdfRect bounds, string? name = null) {
        OdpEllipse shape = OdpEllipse.Create(Presentation, bounds, name ?? NextName("Ellipse")); Element.Add(shape.Element); Dirty(); return shape;
    }
    /// <summary>Adds a text box to the group.</summary>
    public OdpTextBox AddTextBox(OdfRect bounds, string? text = null, string? name = null) {
        OdpTextBox shape = OdpTextBox.Create(Presentation, bounds, text, name ?? NextName("TextBox")); Element.Add(shape.Element); Dirty(); return shape;
    }
    internal static OdpGroup Create(OdpPresentation presentation, string name) => new OdpGroup(presentation,
        new XElement(OdfNamespaces.Draw + "g", new XAttribute(OdfNamespaces.Draw + "name", name)));
    private string NextName(string prefix) => prefix + (Shapes.Count + 1).ToString(CultureInfo.InvariantCulture);
}
