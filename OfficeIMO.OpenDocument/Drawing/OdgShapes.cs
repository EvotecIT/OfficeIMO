using System.Collections;

namespace OfficeIMO.OpenDocument;

/// <summary>An editable drawing-page, master or group collection. Unknown drawing elements remain visible and preserved.</summary>
public sealed partial class OdgShapes : IReadOnlyList<OdgShape> {
    private readonly OdgDocument _document;
    private readonly XElement _parent;
    internal OdgShapes(OdgDocument document, XElement parent) { _document = document; _parent = parent; }
    private string PartPath => _document.GetPartPath(_parent);
    private XElement? DrawingContext => _parent.AncestorsAndSelf().FirstOrDefault(element =>
        element.Name == OdfNamespaces.Draw + "page" || element.Name == OdfNamespaces.Style + "master-page");
    private IEnumerable<XElement> NativeElements => _parent.Elements().Where(element =>
        (element.Name.Namespace == OdfNamespaces.Draw && element.Name.LocalName != "page-thumbnail" && element.Name.LocalName != "layer-set" && element.Name.LocalName != "glue-point") ||
        element.Name.NamespaceName == "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0");
    private IEnumerable<XElement> Elements => _parent.Name == OdfNamespaces.Draw + "g" ? NativeElements
        : NativeElements.Select((element, ordinal) => new { Element = element, Order = ReadZIndex(element, ordinal) })
            .OrderBy(item => item.Order).Select(item => item.Element);
    private static long ReadZIndex(XElement element, int ordinal) =>
        long.TryParse((string?)element.Attribute(OdfNamespaces.Draw + "z-index"), NumberStyles.Integer, CultureInfo.InvariantCulture, out long index) && index >= 0 ? index : ordinal;
    /// <summary>Number of shapes.</summary>
    public int Count => Elements.Count();
    /// <summary>Shape at a zero-based paint-order position.</summary>
    public OdgShape this[int index] => new OdgShape(_document, Elements.ElementAtOrDefault(index) ?? throw new ArgumentOutOfRangeException(nameof(index)));
    /// <summary>Adds a rectangle.</summary>
    public OdgShape AddRectangle(OdfRect bounds, string? name = null) => AddBounded("rect", bounds, name);
    /// <summary>Adds an ellipse.</summary>
    public OdgShape AddEllipse(OdfRect bounds, string? name = null) => AddBounded("ellipse", bounds, name);
    /// <summary>Adds a text frame.</summary>
    public OdgShape AddTextBox(OdfRect bounds, string text, string? name = null) {
        var element = NewElement("frame", name); element.Add(new XElement(OdfNamespaces.Draw + "text-box"));
        OdfShape.ApplyBounds(element, bounds);
        OdgShape shape = Append(element); shape.Text = text; shape.FillColor = null; shape.StrokeColor = null; return shape;
    }
    /// <summary>Adds an embedded image using the existing ODF image store.</summary>
    public OdgShape AddImage(byte[] data, string fileName, OdfRect bounds, string? name = null) {
        string path = OdfImageStore.Add(_document, data, fileName);
        var element = NewElement("frame", name); OdfShape.ApplyBounds(element, bounds);
        element.Add(new XElement(OdfNamespaces.Draw + "image", new XAttribute(OdfNamespaces.XLink + "href", path),
            new XAttribute(OdfNamespaces.XLink + "type", "simple"), new XAttribute(OdfNamespaces.XLink + "show", "embed"), new XAttribute(OdfNamespaces.XLink + "actuate", "onLoad")));
        return Append(element);
    }
    /// <summary>Adds a line with absolute page coordinates.</summary>
    public OdgShape AddLine(OdfLength x1, OdfLength y1, OdfLength x2, OdfLength y2, string? name = null) {
        var element = NewElement("line", name);
        element.SetAttributeValue(OdfNamespaces.Svg + "x1", x1); element.SetAttributeValue(OdfNamespaces.Svg + "y1", y1);
        element.SetAttributeValue(OdfNamespaces.Svg + "x2", x2); element.SetAttributeValue(OdfNamespaces.Svg + "y2", y2);
        OdgShape shape = Append(element); shape.StrokeColor = OdfColor.Parse("000000"); shape.StrokeWidth = OdfLength.Points(1); return shape;
    }
    /// <summary>Adds a straight connector attached to two explicit glue points on this page or master.</summary>
    public OdgShape AddConnector(OdgGluePoint start, OdgGluePoint end, string? name = null) {
        if (start == null) throw new ArgumentNullException(nameof(start));
        if (end == null) throw new ArgumentNullException(nameof(end));
        XElement? page = DrawingContext;
        foreach (OdgGluePoint point in new[] { start, end }) {
            if (!ReferenceEquals(_document, point.Shape.Document) || page == null || !ReferenceEquals(page, point.Shape.PageElement) || point.Element.Parent != point.Shape.Element)
                throw new ArgumentException("Connector endpoints must be live glue points on the same drawing page or master.");
            _ = point.Position;
        }
        var element = NewElement("connector", name);
        element.SetAttributeValue(OdfNamespaces.Svg + "viewBox", "0 0 1 1");
        OdgShape shape = Append(element);
        try {
            shape.ConnectorKind = OdgConnectorKind.Line;
            shape.AttachStart(start); shape.AttachEnd(end);
            shape.StrokeColor = OdfColor.Parse("000000"); shape.StrokeWidth = OdfLength.Points(1);
            return shape;
        } catch { element.Remove(); throw; }
    }
    /// <summary>Adds a group; its Children collection accepts the same shape operations.</summary>
    public OdgShape AddGroup(string? name = null) => Append(NewElement("g", name));
    /// <summary>Removes one shape, including its children.</summary>
    public void RemoveAt(int index) {
        XElement removed = this[index].Element;
        var ids = new HashSet<string>(removed.DescendantsAndSelf().Attributes().Where(attribute =>
            attribute.Name == XNamespace.Xml + "id" || attribute.Name == OdfNamespaces.Draw + "id").Select(attribute => attribute.Value), StringComparer.Ordinal);
        XElement? page = DrawingContext;
        var detachments = new List<Action>();
        foreach (XElement element in page?.Descendants(OdfNamespaces.Draw + "connector") ?? Enumerable.Empty<XElement>()) {
            if (element.AncestorsAndSelf().Contains(removed)) continue;
            var connector = new OdgShape(_document, element);
            foreach (bool start in new[] { true, false }) {
                if (!ids.Contains((start ? connector.StartShapeId : connector.EndShapeId) ?? "")) continue;
                string end = start ? "start" : "end", suffix = start ? "1" : "2";
                var position = connector.DetachmentPosition(start);
                OdfLength x = OdfLength.Points(position.X), y = OdfLength.Points(position.Y);
                detachments.Add(() => {
                    element.SetAttributeValue(OdfNamespaces.Draw + end + "-shape", null);
                    element.SetAttributeValue(OdfNamespaces.Draw + end + "-glue-point", null);
                    element.SetAttributeValue(OdfNamespaces.Svg + "x" + suffix, x);
                    element.SetAttributeValue(OdfNamespaces.Svg + "y" + suffix, y);
                    element.Attribute(OdfNamespaces.Svg + "d")?.Remove();
                });
            }
        }
        foreach (Action detach in detachments) detach();
        removed.Remove(); _document.MarkPartDirty(PartPath);
    }
    /// <summary>Changes a shape's paint-order position.</summary>
    public void Move(int sourceIndex, int destinationIndex) {
        var elements = Elements.ToList();
        if (sourceIndex < 0 || sourceIndex >= elements.Count) throw new ArgumentOutOfRangeException(nameof(sourceIndex));
        if (destinationIndex < 0 || destinationIndex >= elements.Count) throw new ArgumentOutOfRangeException(nameof(destinationIndex));
        bool indexed = elements.Any(element => element.Attribute(OdfNamespaces.Draw + "z-index") != null);
        XElement element = elements[sourceIndex]; element.Remove(); elements.RemoveAt(sourceIndex);
        if (destinationIndex == elements.Count) _parent.Add(element); else elements[destinationIndex].AddBeforeSelf(element);
        elements.Insert(destinationIndex, element);
        if (indexed) UpdateIndices(elements);
        _document.MarkPartDirty(PartPath);
    }
    /// <inheritdoc />
    public IEnumerator<OdgShape> GetEnumerator() => Elements.Select(element => new OdgShape(_document, element)).GetEnumerator();
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    private XElement NewElement(string localName, string? name) => new XElement(OdfNamespaces.Draw + localName,
        new XAttribute(OdfNamespaces.Draw + "name", name ?? OdgDocument.NextName(
            new[] { "content.xml", "styles.xml" }.Where(_document.Package.ContainsEntry)
                .SelectMany(part => _document.GetXml(part).Descendants())
                .Select(element => (string?)element.Attribute(OdfNamespaces.Draw + "name") ?? ""), "Shape")));
    private OdgShape Append(XElement element) {
        var ordered = Elements.ToList();
        bool indexed = ordered.Any(shape => shape.Attribute(OdfNamespaces.Draw + "z-index") != null);
        _parent.Add(element); ordered.Add(element);
        if (indexed) UpdateIndices(ordered);
        _document.MarkPartDirty(PartPath); return new OdgShape(_document, element);
    }
    private void UpdateIndices(IReadOnlyList<XElement> elements) {
        for (int index = 0; index < elements.Count; index++)
            elements[index].SetAttributeValue(OdfNamespaces.Draw + "z-index", _parent.Name == OdfNamespaces.Draw + "g" ? null : index.ToString(CultureInfo.InvariantCulture));
    }
    private OdgShape AddBounded(string localName, OdfRect bounds, string? name) {
        var element = NewElement(localName, name); OdfShape.ApplyBounds(element, bounds);
        OdgShape shape = Append(element); shape.FillColor = OdfColor.Parse("FFFFFF"); shape.StrokeColor = OdfColor.Parse("000000"); shape.StrokeWidth = OdfLength.Points(1); return shape;
    }
}
