namespace OfficeIMO.OpenDocument;

/// <summary>An XML-backed ODP slide.</summary>
public sealed class OdpSlide {
    private readonly OdpPresentation _presentation;
    internal OdpSlide(OdpPresentation presentation, XElement element) { _presentation = presentation; Element = element; }

    /// <summary>Slide name.</summary>
    public string Name {
        get => (string?)Element.Attribute(OdfNamespaces.Draw + "name") ?? string.Empty;
        set { if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("Slide name cannot be empty.", nameof(value)); Element.SetAttributeValue(OdfNamespaces.Draw + "name", value); Dirty(); }
    }
    /// <summary>Whether the slide is hidden from the normal show.</summary>
    public bool Hidden {
        get => (string?)Element.Attribute(OdfNamespaces.Presentation + "visibility") == "hidden";
        set { Element.SetAttributeValue(OdfNamespaces.Presentation + "visibility", value ? "hidden" : null); Dirty(); }
    }
    /// <summary>Referenced master page name.</summary>
    public string? MasterPageName {
        get => (string?)Element.Attribute(OdfNamespaces.Draw + "master-page-name");
        set { Element.SetAttributeValue(OdfNamespaces.Draw + "master-page-name", value); Dirty(); }
    }
    /// <summary>Referenced presentation layout name.</summary>
    public string? LayoutName {
        get => (string?)Element.Attribute(OdfNamespaces.Presentation + "presentation-page-layout-name");
        set { Element.SetAttributeValue(OdfNamespaces.Presentation + "presentation-page-layout-name", value); Dirty(); }
    }
    /// <summary>Slide shapes in XML order.</summary>
    public IReadOnlyList<OdpShape> Shapes => Element.Elements()
        .Where(element => element.Name.Namespace == OdfNamespaces.Draw && element.Name != OdfNamespaces.Draw + "page-thumbnail")
        .Select(element => OdpShape.Wrap(_presentation, element)).Where(shape => shape != null).Select(shape => shape!).ToList();
    /// <summary>Basic typed animation effects in timing-tree order.</summary>
    public IReadOnlyList<OdpAnimation> Animations => Element.Descendants(OdfNamespaces.Anim + "animate")
        .Select(element => new OdpAnimation(_presentation, element)).ToList();
    /// <summary>Speaker notes, or null when none are present.</summary>
    public OdpNotes? SpeakerNotes {
        get {
            XElement? notes = Element.Element(OdfNamespaces.Presentation + "notes");
            return notes == null ? null : new OdpNotes(_presentation, notes);
        }
    }
    /// <summary>Slide background fill color.</summary>
    public OdfColor? BackgroundColor {
        get {
            OdfStyle? style = GetDrawingPageStyle();
            string? value = (string?)style?.Element.Element(OdfNamespaces.Style + "drawing-page-properties")?.Attribute(OdfNamespaces.Draw + "fill-color");
            return value == null ? (OdfColor?)null : OdfColor.Parse(value);
        }
        set {
            OdfStyle style = EnsureDrawingPageStyle();
            style.SetProperty(OdfNamespaces.Style + "drawing-page-properties", OdfNamespaces.Draw + "fill", value.HasValue ? "solid" : null);
            style.SetProperty(OdfNamespaces.Style + "drawing-page-properties", OdfNamespaces.Draw + "fill-color", value?.ToString());
        }
    }
    /// <summary>Raw ODF transition type on the slide's drawing-page style.</summary>
    public string? TransitionType {
        get => (string?)GetDrawingPageStyle()?.Element.Element(OdfNamespaces.Style + "drawing-page-properties")?.Attribute(OdfNamespaces.Presentation + "transition-type");
        set => EnsureDrawingPageStyle().SetProperty(OdfNamespaces.Style + "drawing-page-properties", OdfNamespaces.Presentation + "transition-type", value);
    }
    /// <summary>Raw ODF transition style on the slide's drawing-page style.</summary>
    public string? TransitionStyle {
        get => (string?)GetDrawingPageStyle()?.Element.Element(OdfNamespaces.Style + "drawing-page-properties")?.Attribute(OdfNamespaces.Presentation + "transition-style");
        set => EnsureDrawingPageStyle().SetProperty(OdfNamespaces.Style + "drawing-page-properties", OdfNamespaces.Presentation + "transition-style", value);
    }

    /// <summary>Adds a text box.</summary>
    public OdpTextBox AddTextBox(OdfRect bounds, string? text = null, string? name = null) {
        OdpTextBox box = OdpTextBox.Create(_presentation, bounds, text, name ?? NextShapeName("TextBox")); AddShapeElement(box.Element); Dirty(); return box;
    }
    /// <summary>Adds a rectangle.</summary>
    public OdpRectangle AddRectangle(OdfRect bounds, string? name = null) {
        OdpRectangle shape = OdpRectangle.Create(_presentation, bounds, name ?? NextShapeName("Rectangle")); AddShapeElement(shape.Element); Dirty(); return shape;
    }
    /// <summary>Adds an ellipse.</summary>
    public OdpEllipse AddEllipse(OdfRect bounds, string? name = null) {
        OdpEllipse shape = OdpEllipse.Create(_presentation, bounds, name ?? NextShapeName("Ellipse")); AddShapeElement(shape.Element); Dirty(); return shape;
    }
    /// <summary>Adds a line.</summary>
    public OdpLine AddLine(OdfLength x1, OdfLength y1, OdfLength x2, OdfLength y2, string? name = null) {
        OdpLine shape = OdpLine.Create(_presentation, x1, y1, x2, y2, name ?? NextShapeName("Line")); AddShapeElement(shape.Element); Dirty(); return shape;
    }
    /// <summary>Adds a group.</summary>
    public OdpGroup AddGroup(string? name = null) {
        OdpGroup group = OdpGroup.Create(_presentation, name ?? NextShapeName("Group")); AddShapeElement(group.Element); Dirty(); return group;
    }
    /// <summary>Adds an image frame.</summary>
    public OdpImage AddImage(byte[] data, string fileName, OdfRect bounds, string? name = null) {
        OdpImage image = OdpImage.Create(_presentation, data, fileName, bounds, name ?? NextShapeName("Image")); AddShapeElement(image.Element); Dirty(); return image;
    }
    /// <summary>Adds a presentation table frame.</summary>
    public OdpTable AddTable(OdfRect bounds, int rows, int columns, string? name = null) {
        OdpTable table = OdpTable.Create(_presentation, bounds, rows, columns, name ?? NextShapeName("Table")); AddShapeElement(table.Element); Dirty(); return table;
    }
    /// <summary>Adds a bounded basic animation effect for a shape attribute.</summary>
    public OdpAnimation AddAnimation(OdpShape target, string attributeName, string from, string to, TimeSpan duration) {
        if (target == null) throw new ArgumentNullException(nameof(target));
        if (!ReferenceEquals(target.Presentation, _presentation)) throw new ArgumentException("Animation target must belong to this presentation.", nameof(target));
        if (!Element.DescendantsAndSelf().Contains(target.Element)) throw new ArgumentException("Animation target must belong to this slide.", nameof(target));
        if (string.IsNullOrWhiteSpace(attributeName)) throw new ArgumentException("Animated attribute cannot be empty.", nameof(attributeName));
        if (from == null) throw new ArgumentNullException(nameof(from));
        if (to == null) throw new ArgumentNullException(nameof(to));
        if (duration <= TimeSpan.Zero) throw new ArgumentOutOfRangeException(nameof(duration));
        string targetId = target.EnsureXmlId();
        XElement sequence = EnsureAnimationSequence();
        var element = new XElement(OdfNamespaces.Anim + "animate",
            new XAttribute(OdfNamespaces.Smil + "targetElement", targetId),
            new XAttribute(OdfNamespaces.Smil + "attributeName", attributeName),
            new XAttribute(OdfNamespaces.Smil + "from", from),
            new XAttribute(OdfNamespaces.Smil + "to", to),
            new XAttribute(OdfNamespaces.Smil + "dur", XmlConvert.ToString(duration)),
            new XAttribute(OdfNamespaces.Smil + "fill", "freeze"));
        sequence.Add(new XElement(OdfNamespaces.Anim + "par",
            new XAttribute(OdfNamespaces.Presentation + "node-type", "on-click"), element));
        Dirty();
        return new OdpAnimation(_presentation, element);
    }
    /// <summary>Adds an opacity fade-in effect.</summary>
    public OdpAnimation AddFadeInAnimation(OdpShape target, TimeSpan duration) => AddAnimation(target, "opacity", "0", "1", duration);
    /// <summary>Gets or creates the speaker-notes container.</summary>
    public OdpNotes GetOrCreateSpeakerNotes() {
        XElement? notes = Element.Element(OdfNamespaces.Presentation + "notes");
        if (notes == null) { notes = new XElement(OdfNamespaces.Presentation + "notes"); Element.Add(notes); Dirty(); }
        return new OdpNotes(_presentation, notes);
    }

    internal XElement Element { get; }
    private OdfStyle EnsureDrawingPageStyle() => _presentation.Styles.EnsureAutomaticStyle(Element,
        OdfNamespaces.Draw + "style-name", OdfStyleFamily.DrawingPage, "ofSlide");
    private OdfStyle? GetDrawingPageStyle() {
        string? name = (string?)Element.Attribute(OdfNamespaces.Draw + "style-name");
        return name == null ? null : _presentation.Styles.Find(OdfStyleFamily.DrawingPage, name);
    }
    private XElement EnsureAnimationSequence() {
        XElement? root = Element.Elements(OdfNamespaces.Anim + "par")
            .FirstOrDefault(element => (string?)element.Attribute(OdfNamespaces.Presentation + "node-type") == "timing-root");
        if (root == null) {
            root = new XElement(OdfNamespaces.Anim + "par",
                new XAttribute(OdfNamespaces.Presentation + "node-type", "timing-root"));
            XElement? notes = Element.Element(OdfNamespaces.Presentation + "notes");
            if (notes == null) Element.Add(root); else notes.AddBeforeSelf(root);
        }
        XElement? sequence = root.Element(OdfNamespaces.Anim + "seq");
        if (sequence == null) {
            sequence = new XElement(OdfNamespaces.Anim + "seq",
                new XAttribute(OdfNamespaces.Presentation + "node-type", "main-sequence"));
            root.Add(sequence);
        }
        return sequence;
    }
    private void AddShapeElement(XElement shape) {
        XElement? boundary = Element.Elements().FirstOrDefault(element =>
            element.Name.Namespace == OdfNamespaces.Anim || element.Name == OdfNamespaces.Presentation + "notes");
        if (boundary == null) Element.Add(shape); else boundary.AddBeforeSelf(shape);
    }
    private string NextShapeName(string prefix) {
        var names = new HashSet<string>(Element.Descendants().Select(item => (string?)item.Attribute(OdfNamespaces.Draw + "name")).Where(value => value != null)!, StringComparer.Ordinal);
        int index = 1; string name; do { name = prefix + index++.ToString(CultureInfo.InvariantCulture); } while (names.Contains(name)); return name;
    }
    private void Dirty() => _presentation.MarkPartDirty("content.xml");
}
