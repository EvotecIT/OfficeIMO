using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

/// <summary>Shared XML-backed shape properties for OpenDocument presentations and drawings.</summary>
public abstract partial class OdfShape {
    internal OdfShape(OdfDocument document, XElement element) { Document = document; Element = element; }
    /// <summary>Shape name.</summary>
    public string Name {
        get => (string?)Element.Attribute(OdfNamespaces.Draw + "name") ?? string.Empty;
        set { Element.SetAttributeValue(OdfNamespaces.Draw + "name", value); Dirty(); }
    }
    /// <summary>
    /// XML identifier used by animation and cross-reference targets. Assigned values are unique across the main XML parts.
    /// Connector and animation references in the same page or master follow an identifier change.
    /// </summary>
    public string? XmlId {
        get => (string?)Element.Attribute(XNamespace.Xml + "id") ?? (string?)Element.Attribute(OdfNamespaces.Draw + "id");
        set {
            if (value != null) {
                XmlConvert.VerifyNCName(value);
                bool duplicate = ShapeIdentifierElements()
                    .Any(element => !ReferenceEquals(element, Element) &&
                        ShapeIdentifierAttributes(element).Any(attribute => string.Equals(attribute.Value, value, StringComparison.Ordinal)));
                if (duplicate) throw new ArgumentException("Shape XML identifiers must be unique within the document.", nameof(value));
            }
            string? old = XmlId;
            var references = old == null ? new List<XAttribute>() : LocalShapeReferences(old).ToList();
            if (value == null && references.Count > 0) throw new InvalidOperationException("Remove references before removing a shape identifier.");
            Element.SetAttributeValue(XNamespace.Xml + "id", value);
            if (Element.Attribute(OdfNamespaces.Draw + "id") != null) Element.SetAttributeValue(OdfNamespaces.Draw + "id", value);
            foreach (XAttribute reference in references) reference.Value = value!;
            foreach (string part in references.Select(reference => Document.GetPartPath(reference.Parent!)).Distinct(StringComparer.Ordinal))
                Document.MarkPartDirty(part);
            Dirty();
        }
    }
    /// <summary>Raw ODF/SVG transform expression.</summary>
    public virtual string? Transform {
        get => (string?)Element.Attribute(OdfNamespaces.Draw + "transform");
        set { Element.SetAttributeValue(OdfNamespaces.Draw + "transform", value); Dirty(); }
    }
    /// <summary>Rule used to fill overlapping contours; ODF follows SVG's nonzero default.</summary>
    public OfficeFillRule FillRule {
        get => ReadGraphicProperty(OdfNamespaces.Svg + "fill-rule") switch {
            null or "nonzero" => OfficeFillRule.NonZero, "evenodd" => OfficeFillRule.EvenOdd,
            _ => throw new InvalidDataException("Unsupported shape fill rule.")
        };
        set {
            if (!Enum.IsDefined(typeof(OfficeFillRule), value)) throw new ArgumentOutOfRangeException(nameof(value));
            EnsureGraphicStyle().SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Svg + "fill-rule", value == OfficeFillRule.EvenOdd ? "evenodd" : "nonzero");
        }
    }
    /// <summary>Solid shape fill color.</summary>
    public OdfColor? FillColor {
        get => ReadGraphicColor(OdfNamespaces.Draw + "fill", OdfNamespaces.Draw + "fill-color");
        set {
            OdfStyle style = EnsureGraphicStyle();
            style.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "fill", value.HasValue ? "solid" : "none");
            style.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "fill-color", value?.ToString());
        }
    }
    /// <summary>Shape stroke color, including dashed strokes.</summary>
    public OdfColor? StrokeColor {
        get => ReadGraphicColor(OdfNamespaces.Draw + "stroke", OdfNamespaces.Svg + "stroke-color");
        set {
            string mode = value.HasValue ? ReadGraphicProperty(OdfNamespaces.Draw + "stroke") == "dash" ? "dash" : "solid" : "none";
            OdfStyle style = EnsureGraphicStyle();
            style.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "stroke", mode);
            style.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Svg + "stroke-color", value?.ToString());
        }
    }
    /// <summary>Shape stroke width.</summary>
    public OdfLength? StrokeWidth {
        get {
            OdfStyle? style = GetGraphicStyle();
            string? value = style == null ? null : Document.Styles.Resolve(style)
                .Select(candidate => (string?)candidate.Element
                    .Element(OdfNamespaces.Style + "graphic-properties")?.Attribute(OdfNamespaces.Svg + "stroke-width"))
                .FirstOrDefault(width => width != null);
            value ??= (string?)GetDefaultGraphicProperties()?.Attribute(OdfNamespaces.Svg + "stroke-width");
            return value == null ? (OdfLength?)null : OdfLength.Parse(value);
        }
        set => EnsureGraphicStyle().SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Svg + "stroke-width", value?.ToString());
    }
    /// <summary>Position and size for shapes exposing SVG bounds.</summary>
    public virtual OdfRect Bounds {
        get => new OdfRect(ReadLength("x"), ReadLength("y"), ReadLength("width"), ReadLength("height"));
        set { ApplyBounds(Element, value); Dirty(); }
    }
    internal OdfDocument Document { get; }
    internal XElement Element { get; }
    internal string PartPath => Document.GetPartPath(Element);
    internal static void ApplyBounds(XElement element, OdfRect bounds) {
        element.SetAttributeValue(OdfNamespaces.Svg + "x", bounds.X.ToString());
        element.SetAttributeValue(OdfNamespaces.Svg + "y", bounds.Y.ToString());
        element.SetAttributeValue(OdfNamespaces.Svg + "width", bounds.Width.ToString());
        element.SetAttributeValue(OdfNamespaces.Svg + "height", bounds.Height.ToString());
    }
    internal void Dirty() => Document.MarkPartDirty(PartPath);
    internal OdfStyle EnsureGraphicStyle() => Document.Styles.EnsureAutomaticStyle(Element, OdfNamespaces.Draw + "style-name", OdfStyleFamily.Graphic, "ofGr", PartPath);
    internal string EnsureXmlId() {
        if (!string.IsNullOrWhiteSpace(XmlId)) return XmlId!;
        var ids = new HashSet<string>(ShapeIdentifierElements().SelectMany(ShapeIdentifierAttributes)
            .Select(attribute => attribute.Value), StringComparer.Ordinal);
        int index = 1; string id;
        do { id = "shape" + index++.ToString(CultureInfo.InvariantCulture); } while (ids.Contains(id));
        XmlId = id; return id;
    }
    private IEnumerable<XElement> ShapeIdentifierElements() {
        // Flat serialization combines both parts, so generated identities must also be safe in that container.
        foreach (string part in new[] { "content.xml", "styles.xml" }) {
            if (!Document.Package.ContainsEntry(part)) continue;
            foreach (XElement element in Document.GetXml(part).Descendants()) yield return element;
        }
    }
    private static IEnumerable<XAttribute> ShapeIdentifierAttributes(XElement element) => element.Attributes().Where(attribute =>
        attribute.Name == XNamespace.Xml + "id" ||
        attribute.Name == OdfNamespaces.Draw + "id" && element.Name != OdfNamespaces.Draw + "glue-point" ||
        attribute.Name == OdfNamespaces.Text + "id" && (element.Name == OdfNamespaces.Text + "p" ||
            element.Name == OdfNamespaces.Text + "h" || element.Name == OdfNamespaces.Draw + "text-box"));
    private IEnumerable<XAttribute> LocalShapeReferences(string id) {
        // Imported package parts can reuse IDs. Attachment resolution belongs to a native page or master, not its overlay.
        XElement scope = Element.Ancestors().FirstOrDefault(parent => parent.Name == OdfNamespaces.Draw + "page" ||
            parent.Name == OdfNamespaces.Style + "master-page") ?? Element.Document?.Root ?? Element;
        return scope.DescendantsAndSelf().Attributes().Where(attribute =>
            (attribute.Name == OdfNamespaces.Smil + "targetElement" || attribute.Parent?.Name == OdfNamespaces.Draw + "connector" &&
             (attribute.Name == OdfNamespaces.Draw + "start-shape" || attribute.Name == OdfNamespaces.Draw + "end-shape")) &&
            string.Equals(attribute.Value, id, StringComparison.Ordinal));
    }
    private OdfStyle? GetGraphicStyle() {
        string? name = (string?)Element.Attribute(OdfNamespaces.Draw + "style-name");
        return name == null ? null : Document.Styles.FindInPart(OdfStyleFamily.Graphic, name, PartPath);
    }
    internal string? ReadGraphicProperty(XName name, XName? shorthand = null) {
        OdfStyle? style = GetGraphicStyle();
        return Document.Styles.ResolveWithDefault(style, OdfStyleFamily.Graphic)
            .Select(candidate => (string?)candidate.Element.Element(OdfNamespaces.Style + "graphic-properties")?.Attribute(name)
                ?? (shorthand == null ? null : (string?)candidate.Element.Element(OdfNamespaces.Style + "graphic-properties")?.Attribute(shorthand)))
            .FirstOrDefault(value => value != null);
    }
    internal string? ReadGraphicWritingMode() => Document.Styles.ResolveWithDefault(GetGraphicStyle(), OdfStyleFamily.Graphic)
        // ODF permits graphic writing mode in either properties element of a graphic style.
        .Select(candidate => (string?)candidate.Element.Element(OdfNamespaces.Style + "graphic-properties")?
            .Attribute(OdfNamespaces.Style + "writing-mode") ?? candidate.WritingMode)
        .FirstOrDefault(value => value != null);
    private OdfColor? ReadGraphicColor(XName modeName, XName colorName) {
        OdfStyle? style = GetGraphicStyle();
        IReadOnlyList<OdfStyle> chain = style == null ? Array.Empty<OdfStyle>() : Document.Styles.Resolve(style);
        string? mode = chain.Select(candidate => (string?)candidate.Element
            .Element(OdfNamespaces.Style + "graphic-properties")?.Attribute(modeName))
            .FirstOrDefault(value => value != null);
        XElement? defaults = GetDefaultGraphicProperties();
        mode ??= (string?)defaults?.Attribute(modeName);
        if (mode != null && mode != "solid" && !(modeName == OdfNamespaces.Draw + "stroke" && mode == "dash")) return null;
        string? value = chain.Select(candidate => (string?)candidate.Element
            .Element(OdfNamespaces.Style + "graphic-properties")?.Attribute(colorName))
            .FirstOrDefault(color => color != null);
        value ??= (string?)defaults?.Attribute(colorName);
        return value == null ? (OdfColor?)null : OdfColor.Parse(value);
    }
    private XElement? GetDefaultGraphicProperties() => Document.Styles.FindDefaultProperties(
        OdfStyleFamily.Graphic, OdfNamespaces.Style + "graphic-properties");
    private OdfLength ReadLength(string localName) => OdfLength.Parse((string?)Element.Attribute(OdfNamespaces.Svg + localName) ?? "0cm");
}
