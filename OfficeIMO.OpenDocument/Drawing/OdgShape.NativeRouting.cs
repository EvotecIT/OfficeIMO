using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

public sealed partial class OdgShape {
    /// <summary>
    /// Encodes the saved three-segment route as native Lines routing, departure directions and segment offsets.
    /// Both ends must attach to untransformed rectangles, frames or full ellipses/circles with declared SVG bounds. The outer segments must be
    /// horizontal or vertical and nonzero; the middle segment may be diagonal. Equivalent directional glue
    /// points are reused, otherwise cloned without altering shared points. The connector's own transform is
    /// baked into the route; transformed parents and transformed connector text are rejected.
    /// Native applications may adjust bends for painted bounds, strokes or subsequent shape edits.
    /// Unsupported input leaves the document unchanged.
    /// </summary>
    public void UseNativeThreeSegmentRouting() {
        RequireConnector();
        if (StartShapeId == null || EndShapeId == null)
            throw new NotSupportedException("Native three-segment routing requires both endpoints to be attached.");
        if (Element.Ancestors(OdfNamespaces.Draw + "g").Any(group =>
            OdfDrawingTransform.Parse((string?)group.Attribute(OdfNamespaces.Draw + "transform")) != OfficeTransform.Identity))
            throw new NotSupportedException("Bake parent transforms before encoding native three-segment routing.");
        OfficeTransform transform = OdfDrawingTransform.Parse(Transform);
        if (transform != OfficeTransform.Identity && HasConnectorLabel)
            throw new NotSupportedException("Baking a connector transform with text is outside the native routing profile.");
        var path = ConnectorRouteCommands.Select(command => MapConnectorCommand(command, transform.TransformPoint)).ToArray();
        if (path.Length != 4 || path.Skip(1).Any(command => command.Kind != OfficePathCommandKind.LineTo))
            throw new NotSupportedException("Native three-segment routing requires one move and three straight segments.");
        OdgGluePointEscapeDirection startDirection = DepartureDirection(path[0].Point, path[1].Point);
        OdgGluePointEscapeDirection endDirection = DepartureDirection(path[3].Point, path[2].Point);
        var additions = new Dictionary<OdgShape, List<XElement>>();
        var start = PrepareNativeAttachment(true, startDirection, path[0].Point, additions);
        var end = PrepareNativeAttachment(false, endDirection, path[3].Point, additions);
        // Snap attached endpoints before rounding, just as SetConnectorRoute does.
        path[0] = OfficePathCommand.MoveTo(new OdgGluePoint(start.Shape, start.Point).Position);
        path[3] = OfficePathCommand.LineTo(new OdgGluePoint(end.Shape, end.Point).Position);
        // Keep the departure axis when endpoint snapping and transform baking cross a native grid boundary.
        path[1] = OfficePathCommand.LineTo(SnapDepartureBend(path[0].Point, path[1].Point, startDirection));
        path[2] = OfficePathCommand.LineTo(SnapDepartureBend(path[3].Point, path[2].Point, endDirection));
        if (NativeConnectorPoint(path[0].Point).Equals(NativeConnectorPoint(path[1].Point)) ||
            NativeConnectorPoint(path[3].Point).Equals(NativeConnectorPoint(path[2].Point)))
            throw new NotSupportedException("Native departure segments must remain nonzero after native-grid rounding.");
        XAttribute[] attributes = EncodeConnectorRoute(path);
        string offsets = NativeSegmentOffset(start.Shape, startDirection, path[1].Point) + " " +
            NativeSegmentOffset(end.Shape, endDirection, path[2].Point);

        OdfStyle style = EnsureGraphicStyle(); // Clones styles shared by other shapes.
        foreach (string endpoint in new[] { "start", "end" }) foreach (string axis in new[] { "horizontal", "vertical" })
            style.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + endpoint + "-line-spacing-" + axis, "0cm");
        foreach (var addition in additions) foreach (XElement point in addition.Value) addition.Key.InsertGluePoint(point);
        Element.SetAttributeValue(OdfNamespaces.Draw + "start-glue-point", (string?)start.Point.Attribute(OdfNamespaces.Draw + "id"));
        Element.SetAttributeValue(OdfNamespaces.Draw + "end-glue-point", (string?)end.Point.Attribute(OdfNamespaces.Draw + "id"));
        foreach (XAttribute attribute in attributes) Element.SetAttributeValue(attribute.Name, attribute.Value);
        Element.SetAttributeValue(OdfNamespaces.Draw + "type", "lines");
        Element.SetAttributeValue(OdfNamespaces.Draw + "line-skew", offsets);
        Element.Attribute(OdfNamespaces.Draw + "transform")?.Remove(); Dirty();
    }

    private static OfficePoint SnapDepartureBend(OfficePoint endpoint, OfficePoint bend, OdgGluePointEscapeDirection direction) =>
        direction is OdgGluePointEscapeDirection.Left or OdgGluePointEscapeDirection.Right
            ? new OfficePoint(bend.X, endpoint.Y) : new OfficePoint(endpoint.X, bend.Y);

    private static OdgGluePointEscapeDirection DepartureDirection(OfficePoint endpoint, OfficePoint bend) {
        double x = bend.X - endpoint.X, y = bend.Y - endpoint.Y;
        if (Math.Abs(x) <= 1e-9 && Math.Abs(y) > 1e-9) return y < 0 ? OdgGluePointEscapeDirection.Up : OdgGluePointEscapeDirection.Down;
        if (Math.Abs(y) <= 1e-9 && Math.Abs(x) > 1e-9) return x < 0 ? OdgGluePointEscapeDirection.Left : OdgGluePointEscapeDirection.Right;
        throw new NotSupportedException("Native departure segments must be nonzero and horizontal or vertical in page coordinates.");
    }

    private (OdgShape Shape, XElement Point) PrepareNativeAttachment(bool start, OdgGluePointEscapeDirection direction,
        OfficePoint endpoint, Dictionary<OdgShape, List<XElement>> additions) {
        string id = (start ? StartShapeId : EndShapeId)!;
        XElement target = PageElement!.Descendants().First(element =>
            (string?)element.Attribute(XNamespace.Xml + "id") == id || (string?)element.Attribute(OdfNamespaces.Draw + "id") == id);
        OdgShape shape = additions.Keys.FirstOrDefault(candidate => ReferenceEquals(candidate.Element, target)) ?? new OdgShape((OdgDocument)Document, target);
        if (shape.ElementName is not ("rect" or "ellipse" or "circle" or "frame") ||
            ((string?)target.Attribute(OdfNamespaces.Draw + "kind") is string kind && kind != "full") ||
            target.Attribute(OdfNamespaces.Svg + "width") == null || target.Attribute(OdfNamespaces.Svg + "height") == null ||
            shape.PageTransform != OfficeTransform.Identity)
            throw new NotSupportedException("Native three-segment attachments require untransformed rectangles, frames or full ellipses/circles with declared SVG bounds.");
        _ = StandardGluePoints(shape).ToArray(); // Validate finite declared bounds.
        string? glueId = (string?)Element.Attribute(OdfNamespaces.Draw + (start ? "start" : "end") + "-glue-point");
        XElement? source = target.Elements(OdfNamespaces.Draw + "glue-point").FirstOrDefault(point => (string?)point.Attribute(OdfNamespaces.Draw + "id") == glueId);
        XElement proposed;
        if (source != null) proposed = new XElement(source);
        else {
            int index = Array.FindIndex(StandardGluePoints(shape).ToArray(), point => Near(point, endpoint));
            if (index < 0) throw new NotSupportedException("The automatic attachment has no matching standard edge center.");
            proposed = new XElement(OdfNamespaces.Draw + "glue-point",
                new XAttribute(OdfNamespaces.Draw + "align", new[] { "top", "right", "bottom", "left" }[index]),
                new XAttribute(OdfNamespaces.Svg + "x", "0cm"), new XAttribute(OdfNamespaces.Svg + "y", "0cm"));
            if (index == 2) {
                proposed.Attribute(OdfNamespaces.Draw + "align")!.Remove();
                proposed.SetAttributeValue(OdfNamespaces.Svg + "x", "0cm"); proposed.SetAttributeValue(OdfNamespaces.Svg + "y", "5cm");
            }
        }
        if (proposed.Attribute(OdfNamespaces.Draw + "align") == null) foreach (string axis in new[] { "x", "y" }) {
            XAttribute? coordinate = proposed.Attribute(OdfNamespaces.Svg + axis);
            if (coordinate != null && coordinate.Value.EndsWith("%", StringComparison.Ordinal)) {
                // LibreOffice imports relative glue coordinates as lengths over a 10,000-unit canvas,
                // and drops percentage tokens. Normalize the new directional copy, preserving its source.
                double percentage = double.Parse(coordinate.Value.Substring(0, coordinate.Value.Length - 1), NumberStyles.Float, CultureInfo.InvariantCulture);
                coordinate.Value = OdfLength.Centimeters(percentage / 10).ToString();
            }
        }
        proposed.Attribute(OdfNamespaces.Draw + "id")?.Remove();
        proposed.SetAttributeValue(OdfNamespaces.Draw + "escape-direction", direction.ToString().ToLowerInvariant());
        if (!additions.TryGetValue(shape, out List<XElement>? pending)) { pending = new List<XElement>(); additions.Add(shape, pending); }
        foreach (XElement point in target.Elements(OdfNamespaces.Draw + "glue-point").Concat(pending)) {
            var comparison = new XElement(point); comparison.Attribute(OdfNamespaces.Draw + "id")?.Remove();
            // Attribute order is immaterial to ODF; compare normalized detached elements.
            if (XNode.DeepEquals(NormalizeGluePoint(comparison), NormalizeGluePoint(proposed))) {
                int existingId = new OdgGluePoint(shape, point).Id;
                if (existingId < 4 || existingId > ushort.MaxValue) throw new NotSupportedException("Custom native glue point IDs must be between 4 and 65535.");
                return (shape, point);
            }
        }
        var ids = new HashSet<int>(target.Elements(OdfNamespaces.Draw + "glue-point").Concat(pending)
            .Select(point => int.Parse((string?)point.Attribute(OdfNamespaces.Draw + "id") ?? throw new InvalidDataException("Missing glue point ID."), CultureInfo.InvariantCulture)));
        int next = 4; while (next <= ushort.MaxValue && ids.Contains(next)) next++;
        if (next > ushort.MaxValue) throw new NotSupportedException("No native glue point identifier is available.");
        proposed.SetAttributeValue(OdfNamespaces.Draw + "id", next); pending.Add(proposed);
        return (shape, proposed);
    }

    private static XElement NormalizeGluePoint(XElement point) => new XElement(point.Name,
        point.Attributes().OrderBy(attribute => attribute.Name.ToString(), StringComparer.Ordinal), point.Nodes());

    private static string NativeSegmentOffset(OdgShape shape, OdgGluePointEscapeDirection direction, OfficePoint bend) {
        OdfRect bounds = shape.Bounds;
        OfficePoint nativeBend = NativeConnectorPoint(bend);
        // Derive offsets from the same rounded bend written to svg:d, so repeated encoding cannot drift.
        double offset = direction switch {
            OdgGluePointEscapeDirection.Left => nativeBend.X - bounds.X.ToPoints() * 2540 / 72,
            OdgGluePointEscapeDirection.Right => nativeBend.X - (bounds.X.ToPoints() + bounds.Width.ToPoints()) * 2540 / 72,
            OdgGluePointEscapeDirection.Up => nativeBend.Y - bounds.Y.ToPoints() * 2540 / 72,
            _ => nativeBend.Y - (bounds.Y.ToPoints() + bounds.Height.ToPoints()) * 2540 / 72
        };
        if (!Finite(offset) || Math.Abs(offset) > int.MaxValue)
            throw new ArgumentOutOfRangeException(nameof(bend), "Native segment offset exceeds signed 32-bit coordinates.");
        return OdfLength.Centimeters(Math.Round(offset, MidpointRounding.AwayFromZero) / 1000).ToString();
    }
}
