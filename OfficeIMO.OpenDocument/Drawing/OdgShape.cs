namespace OfficeIMO.OpenDocument;

/// <summary>A Draw shape backed by its original XML, including elements outside the editable profile.</summary>
public sealed partial class OdgShape : OdfShape {
    internal OdgShape(OdgDocument document, XElement element) : base(document, element) { }
    /// <summary>Native element name, for example rect, ellipse, frame, line, or g.</summary>
    public string ElementName => Element.Name.LocalName;
    /// <summary>Whether this is a shape group.</summary>
    public bool IsGroup => Element.Name == OdfNamespaces.Draw + "g";
    /// <summary>Children of a group. Access on other shapes throws.</summary>
    public OdgShapes Children => IsGroup ? new OdgShapes((OdgDocument)Document, Element) : throw new InvalidOperationException("Only a group has child shapes.");
    /// <summary>Whether the shape has an embedded or linked image.</summary>
    public bool IsImage => Element.Element(OdfNamespaces.Draw + "image") != null;
    /// <summary>Layer name. Layer definitions and visibility remain preserved in the document XML.</summary>
    public string? Layer { get => (string?)Element.Attribute(OdfNamespaces.Draw + "layer"); set { Element.SetAttributeValue(OdfNamespaces.Draw + "layer", value); Dirty(); } }
    /// <summary>Geometry for rectangular shapes. Lines use endpoints; groups use child geometry and transforms.</summary>
    public override OdfRect Bounds {
        get => base.Bounds;
        set {
            if (IsGroup || ElementName is "line" or "connector") throw new NotSupportedException("Use child geometry or line endpoints.");
            base.Bounds = value;
        }
    }
    /// <summary>Plain text of paragraphs, headings and lists in the selected text container, including text boxes and image captions. Assignment replaces their rich formatting.</summary>
    public string Text {
        get => OdfTextCodec.ReadJoined(OdfTextTraversal.Paragraphs(TextRoot));
        set {
            RequireTextEditing();
            TextRoot.Elements().Where(IsTextContainer).Remove();
            foreach (string line in (value ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n').Split('\n')) {
                var paragraph = new XElement(OdfNamespaces.Text + "p"); OdfTextCodec.Append(paragraph, line); OdfDrawTextInsertion.Append(TextRoot, paragraph);
            }
            Dirty();
        }
    }
    /// <summary>Uniform font size for the shape's plain text profile.</summary>
    public OdfLength? FontSize { get => Resolve(style => style.FontSize); set => EnsureGraphicStyle().FontSize = value; }
    /// <summary>Uniform font family for the shape's plain text profile.</summary>
    public string? FontFamily {
        get => TextStyles.Select(style => style.FontFamily).FirstOrDefault(value => value != null);
        set => EnsureGraphicStyle().FontFamily = value;
    }
    /// <summary>Uniform text color for the shape's plain text profile.</summary>
    public OdfColor? TextColor { get => Resolve(style => style.Color); set => EnsureGraphicStyle().Color = value; }
    /// <summary>Line start horizontal coordinate.</summary>
    public OdfLength X1 { get => ReadEndpoint("x1"); set => SetEndpoint("x1", value); }
    /// <summary>Line start vertical coordinate.</summary>
    public OdfLength Y1 { get => ReadEndpoint("y1"); set => SetEndpoint("y1", value); }
    /// <summary>Line end horizontal coordinate.</summary>
    public OdfLength X2 { get => ReadEndpoint("x2"); set => SetEndpoint("x2", value); }
    /// <summary>Line end vertical coordinate.</summary>
    public OdfLength Y2 { get => ReadEndpoint("y2"); set => SetEndpoint("y2", value); }
    /// <summary>Returns embedded image bytes without resolving external image references.</summary>
    public byte[]? GetImageBytes() {
        string? path = (string?)Element.Element(OdfNamespaces.Draw + "image")?.Attribute(OdfNamespaces.XLink + "href");
        if (path != null) path = OdfPackagePath.NormalizeHref(path);
        return path != null && Document.Package.ContainsEntry(path) ? Document.GetPackageEntryBytes(path) : null;
    }
    /// <summary>Returns a detached copy of the native element for inspection.</summary>
    public XElement ToXml() => new XElement(Element);
    internal XElement TextRoot => Element.Element(OdfNamespaces.Draw + "text-box") ?? Element.Element(OdfNamespaces.Draw + "image") ?? Element;
    private static bool IsParagraph(XElement element) => element.Name == OdfNamespaces.Text + "p" || element.Name == OdfNamespaces.Text + "h";
    private static bool IsTextContainer(XElement element) => IsParagraph(element) || element.Name == OdfNamespaces.Text + "list";
    private IEnumerable<OdfStyle> TextStyles {
        get {
            string? name = (string?)Element.Attribute(OdfNamespaces.Draw + "style-name");
            OdfStyle? style = name == null ? null : Document.Styles.FindInPart(OdfStyleFamily.Graphic, name, PartPath);
            return Document.Styles.ResolveWithDefault(style, OdfStyleFamily.Graphic);
        }
    }
    private T? Resolve<T>(Func<OdfStyle, T?> selector) where T : struct => TextStyles.Select(selector).FirstOrDefault(value => value.HasValue);
    private OdfLength ReadEndpoint(string name) {
        if (ElementName is not ("line" or "connector")) throw new InvalidOperationException("Only lines and connectors have editable endpoints.");
        return AttachedCoordinate(name) ?? OdfLength.Parse((string?)Element.Attribute(OdfNamespaces.Svg + name) ?? "0cm");
    }
    private void SetEndpoint(string name, OdfLength value) {
        if (ElementName is not ("line" or "connector")) throw new InvalidOperationException("Only lines and connectors have editable endpoints.");
        if (IsConnector && (name.EndsWith("1", StringComparison.Ordinal) ? StartShapeId : EndShapeId) != null)
            throw new InvalidOperationException("Detach the endpoint before assigning its coordinates.");
        if (!value.TryToPoints(out _)) throw new ArgumentException("Line endpoints must be finite absolute lengths.", nameof(value));
        if (IsConnector) Element.Attribute(OdfNamespaces.Svg + "d")?.Remove();
        Element.SetAttributeValue(OdfNamespaces.Svg + name, value); Dirty();
    }
}
