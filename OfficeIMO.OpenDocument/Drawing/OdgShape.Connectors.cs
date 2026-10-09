using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

/// <summary>Native connector routing style. Saved routes are projected when their endpoints remain consistent.</summary>
public enum OdgConnectorKind {
    /// <summary>A straight segment.</summary>
    Line,
    /// <summary>An orthogonal route.</summary>
    Standard,
    /// <summary>Three connected straight segments.</summary>
    Lines,
    /// <summary>A curved route.</summary>
    Curve
}

public sealed partial class OdgShape {
    /// <summary>Whether this element is a native connector.</summary>
    public bool IsConnector => ElementName == "connector";
    /// <summary>Explicit glue points on this shape.</summary>
    public IReadOnlyList<OdgGluePoint> GluePoints => Element.Elements(OdfNamespaces.Draw + "glue-point").Select(point => new OdgGluePoint(this, point)).ToList();
    /// <summary>
    /// Adds an attachment point with absolute offsets from the selected edge or corner.
    /// Bottom uses a native relative point because ODF has no bottom-center alignment token; its offsets must be zero.
    /// </summary>
    public OdgGluePoint AddGluePoint(OdgGluePointAlignment alignment = OdgGluePointAlignment.Center, OdfLength? offsetX = null, OdfLength? offsetY = null) {
        if (IsGroup || ElementName is "line" or "connector") throw new NotSupportedException("Add a glue point to a bounded shape.");
        if (!Enum.IsDefined(typeof(OdgGluePointAlignment), alignment)) throw new ArgumentOutOfRangeException(nameof(alignment));
        OdfLength x = offsetX ?? OdfLength.Points(0), y = offsetY ?? OdfLength.Points(0);
        if (!x.TryToPoints(out _) || !y.TryToPoints(out _)) throw new ArgumentException("Glue point offsets must be absolute lengths.");
        if (alignment == OdgGluePointAlignment.Bottom && (x.ToPoints() != 0 || y.ToPoints() != 0))
            throw new NotSupportedException("ODF bottom-center relative points require zero offsets.");
        var ids = new HashSet<int>(GluePoints.Select(point => point.Id)); int id = 4; while (ids.Contains(id)) id++;
        string align = alignment switch { OdgGluePointAlignment.TopLeft => "top-left", OdgGluePointAlignment.TopRight => "top-right",
            OdgGluePointAlignment.BottomLeft => "bottom-left", OdgGluePointAlignment.BottomRight => "bottom-right", _ => alignment.ToString().ToLowerInvariant() };
        var element = new XElement(OdfNamespaces.Draw + "glue-point", new XAttribute(OdfNamespaces.Draw + "id", id),
            new XAttribute(OdfNamespaces.Draw + "escape-direction", "auto"), new XAttribute(OdfNamespaces.Draw + "align", align),
            new XAttribute(OdfNamespaces.Svg + "x", x), new XAttribute(OdfNamespaces.Svg + "y", y));
        if (alignment == OdgGluePointAlignment.Bottom) {
            element.Attribute(OdfNamespaces.Draw + "align")!.Remove();
            element.SetAttributeValue(OdfNamespaces.Svg + "x", "0cm"); element.SetAttributeValue(OdfNamespaces.Svg + "y", "5cm");
        }
        InsertGluePoint(element);
        Dirty(); return new OdgGluePoint(this, element);
    }
    private void InsertGluePoint(XElement element) {
        XElement? following = Element.Elements().FirstOrDefault(child => ElementName == "frame"
            ? child.Name == OdfNamespaces.Draw + "image-map" || child.Name.Namespace == OdfNamespaces.Svg || child.Name.LocalName.StartsWith("contour-", StringComparison.Ordinal)
            : child.Name.Namespace == OdfNamespaces.Text || child.Name == OdfNamespaces.Draw + "enhanced-geometry");
        if (following == null) Element.Add(element); else following.AddBeforeSelf(element);
    }
    /// <summary>Native connector routing style.</summary>
    public OdgConnectorKind ConnectorKind {
        get {
            RequireConnector();
            return ((string?)Element.Attribute(OdfNamespaces.Draw + "type")) switch {
                null or "standard" => OdgConnectorKind.Standard, "line" => OdgConnectorKind.Line, "lines" => OdgConnectorKind.Lines, "curve" => OdgConnectorKind.Curve,
                _ => throw new InvalidDataException("Invalid connector type.")
            };
        }
        set {
            RequireConnector();
            if (!Enum.IsDefined(typeof(OdgConnectorKind), value)) throw new ArgumentOutOfRangeException(nameof(value));
            if (ConnectorKind != value) Element.Attribute(OdfNamespaces.Svg + "d")?.Remove();
            Element.SetAttributeValue(OdfNamespaces.Draw + "type", value.ToString().ToLowerInvariant()); Dirty();
        }
    }
    /// <summary>Identifier of the shape attached to the start, or null for a free endpoint.</summary>
    public string? StartShapeId => (string?)Element.Attribute(OdfNamespaces.Draw + "start-shape");
    /// <summary>Identifier of the shape attached to the end, or null for a free endpoint.</summary>
    public string? EndShapeId => (string?)Element.Attribute(OdfNamespaces.Draw + "end-shape");
    /// <summary>Attaches the start to a glue point. Null detaches it while retaining its current position.</summary>
    public void AttachStart(OdgGluePoint? point) => Attach(point, true);
    /// <summary>Attaches the end to a glue point. Null detaches it while retaining its current position.</summary>
    public void AttachEnd(OdgGluePoint? point) => Attach(point, false);
    /// <summary>Attaches the start using automatic selection among the shape's four edge centers. Null detaches it.</summary>
    public void AttachStartToShape(OdgShape? shape) => AttachToShape(shape, true);
    /// <summary>Attaches the end using automatic selection among the shape's four edge centers. Null detaches it.</summary>
    public void AttachEndToShape(OdgShape? shape) => AttachToShape(shape, false);
    private void AttachToShape(OdgShape? shape, bool start) {
        RequireConnector();
        if (shape == null) { Attach(null, start); return; }
        ValidateAttachmentShape(shape);
        _ = InversePageTransform();
        string end = start ? "start" : "end", suffix = start ? "1" : "2";
        Element.SetAttributeValue(OdfNamespaces.Draw + end + "-shape", shape.EnsureXmlId());
        Element.Attribute(OdfNamespaces.Draw + end + "-glue-point")?.Remove();
        Element.Attribute(OdfNamespaces.Svg + "x" + suffix)?.Remove();
        Element.Attribute(OdfNamespaces.Svg + "y" + suffix)?.Remove();
        Element.Attribute(OdfNamespaces.Svg + "d")?.Remove(); Dirty();
    }
    private void ValidateAttachmentShape(OdgShape shape) {
        if (!ReferenceEquals(Document, shape.Document) || PageElement == null || !ReferenceEquals(PageElement, shape.PageElement))
            throw new ArgumentException("Connector endpoints must belong to the same drawing page.", nameof(shape));
        if (shape.IsGroup || shape.ElementName is "line" or "connector") throw new NotSupportedException("Attachment to an unbounded shape is not projected.");
        _ = StandardGluePoints(shape).ToArray();
    }
    private void RequireConnector() { if (!IsConnector) throw new InvalidOperationException("Only connectors have attachments and routing."); }
    private void Attach(OdgGluePoint? point, bool start) {
        RequireConnector();
        if (point != null && (!ReferenceEquals(Document, point.Shape.Document) || PageElement == null || !ReferenceEquals(PageElement, point.Shape.PageElement)))
            throw new ArgumentException("Connector endpoints must belong to the same drawing page.", nameof(point));
        if (point != null && point.Element.Parent != point.Shape.Element) throw new ArgumentException("The glue point is no longer attached to its shape.", nameof(point));
        string end = start ? "start" : "end", suffix = start ? "1" : "2";
        OfficePoint location = point == null ? DetachmentPosition(start)
            : InversePageTransform().TransformPoint(point.Position);
        Element.SetAttributeValue(OdfNamespaces.Draw + end + "-shape", point?.Shape.EnsureXmlId());
        Element.SetAttributeValue(OdfNamespaces.Draw + end + "-glue-point", point?.Id);
        // Native applications calculate attached endpoint coordinates from the glue point.
        Element.SetAttributeValue(OdfNamespaces.Svg + "x" + suffix, point == null ? OdfLength.Points(location.X).ToString() : null);
        Element.SetAttributeValue(OdfNamespaces.Svg + "y" + suffix, point == null ? OdfLength.Points(location.Y).ToString() : null);
        Element.Attribute(OdfNamespaces.Svg + "d")?.Remove(); Dirty();
    }
    internal OfficePoint DetachmentPosition(bool start) {
        string suffix = start ? "1" : "2";
        try { return new OfficePoint(ReadEndpoint("x" + suffix).ToPoints(), ReadEndpoint("y" + suffix).ToPoints()); }
        catch (NotSupportedException) {
            // Native routing can be outside projection while its saved endpoint is still usable for editing.
            string? x = (string?)Element.Attribute(OdfNamespaces.Svg + "x" + suffix), y = (string?)Element.Attribute(OdfNamespaces.Svg + "y" + suffix);
            if (x == null || y == null) throw;
            return new OfficePoint(OdfLength.Parse(x).ToPoints(), OdfLength.Parse(y).ToPoints());
        }
    }
    internal XElement? PageElement => Element.Ancestors().FirstOrDefault(parent =>
        parent.Name == OdfNamespaces.Draw + "page" || parent.Name == OdfNamespaces.Style + "master-page");
    internal OfficeTransform PageTransform {
        get {
            OfficeTransform result = OdfDrawingTransform.Parse(Transform);
            foreach (XElement parent in Element.Ancestors(OdfNamespaces.Draw + "g")) result = result.Then(OdfDrawingTransform.Parse((string?)parent.Attribute(OdfNamespaces.Draw + "transform")));
            return result;
        }
    }
    private OfficeTransform InversePageTransform() => PageTransform.TryInvert(out OfficeTransform inverse) ? inverse
        : throw new NotSupportedException("Connector transform is singular and cannot resolve attached endpoints.");
    private OdfLength? AttachedCoordinate(string name) {
        if (!IsConnector) return null;
        bool start = name.EndsWith("1", StringComparison.Ordinal);
        string end = start ? "start" : "end";
        string? shapeId = start ? StartShapeId : EndShapeId;
        if (shapeId == null) return null;
        XElement? target = PageElement?.Descendants().FirstOrDefault(element =>
            (string?)element.Attribute(XNamespace.Xml + "id") == shapeId || (string?)element.Attribute(OdfNamespaces.Draw + "id") == shapeId);
        if (target == null) throw new InvalidDataException("Connector target is missing: " + shapeId);
        if (target.Name.LocalName is "connector" or "line" or "g") throw new NotSupportedException("Attachment to an unbounded shape is not projected.");
        string? glueId = (string?)Element.Attribute(OdfNamespaces.Draw + end + "-glue-point");
        XElement? glue = target.Elements(OdfNamespaces.Draw + "glue-point").FirstOrDefault(element => (string?)element.Attribute(OdfNamespaces.Draw + "id") == glueId);
        var shape = AttachmentShape(target);
        OfficePoint position;
        if (glue != null) position = new OdgGluePoint(shape, glue).Position;
        else if (glueId is "0" or "1" or "2" or "3") position = StandardGluePoints(shape).ElementAt(int.Parse(glueId, CultureInfo.InvariantCulture));
        else if (glueId == null) position = AutomaticGluePoint(shape, start);
        else throw new NotSupportedException("Connector glue point is missing: " + glueId);
        OfficePoint local = InversePageTransform().TransformPoint(position);
        return OdfLength.Points(name[0] == 'x' ? local.X : local.Y);
    }
}
