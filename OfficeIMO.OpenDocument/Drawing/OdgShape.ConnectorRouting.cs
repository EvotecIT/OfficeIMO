using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

public sealed partial class OdgShape {
    // Native producers round geometry and saved glue positions to 1/100 mm.
    internal const double ConnectorEndpointTolerance = 0.1;
    internal bool UsesAutomaticGlue => IsConnector &&
        ((StartShapeId != null && Element.Attribute(OdfNamespaces.Draw + "start-glue-point") == null) ||
         (EndShapeId != null && Element.Attribute(OdfNamespaces.Draw + "end-glue-point") == null));

    private static IEnumerable<OfficePoint> StandardGluePoints(OdgShape shape) {
        OdfRect bounds = shape.AttachmentBounds;
        double x = bounds.X.ToPoints(), y = bounds.Y.ToPoints(), w = bounds.Width.ToPoints(), h = bounds.Height.ToPoints();
        if (!Finite(x) || !Finite(y) || !Finite(w) || !Finite(h) || w < 0 || h < 0)
            throw new InvalidDataException("Attachment bounds must be finite and nonnegative.");
        OfficeTransform transform = shape.PageTransform;
        yield return transform.TransformPoint(new OfficePoint(x + w / 2, y));
        yield return transform.TransformPoint(new OfficePoint(x + w, y + h / 2));
        yield return transform.TransformPoint(new OfficePoint(x + w / 2, y + h));
        yield return transform.TransformPoint(new OfficePoint(x, y + h / 2));
    }

    private OfficePoint AutomaticGluePoint(OdgShape target, bool start) {
        var points = StandardGluePoints(target).ToArray();
        if (Element.Attribute(OdfNamespaces.Svg + "d") is XAttribute data) {
            var path = ParsePath(data.Value); ValidateConnectorPath(path);
            OfficePoint routePoint = (start ? path[0].Point : path[path.Count - 1].Point);
            routePoint = PageTransform.TransformPoint(new OfficePoint(routePoint.X * 72 / 2540, routePoint.Y * 72 / 2540));
            foreach (OfficePoint point in points) if (Near(point, routePoint)) return point;
        }
        if (TryReadSavedEndpoint(start, out OfficePoint saved)) {
            OfficePoint pagePoint = PageTransform.TransformPoint(saved);
            foreach (OfficePoint point in points) if (Near(point, pagePoint)) return point;
        }
        string? otherId = start ? EndShapeId : StartShapeId;
        OfficePoint other;
        if (otherId == null) {
            other = PageTransform.TransformPoint(new OfficePoint(
                OdfLength.Parse((string?)Element.Attribute(OdfNamespaces.Svg + (start ? "x2" : "x1")) ?? "0cm").ToPoints(),
                OdfLength.Parse((string?)Element.Attribute(OdfNamespaces.Svg + (start ? "y2" : "y1")) ?? "0cm").ToPoints()));
        } else {
            XElement? element = PageElement?.Descendants().FirstOrDefault(candidate =>
                (string?)candidate.Attribute(XNamespace.Xml + "id") == otherId || (string?)candidate.Attribute(OdfNamespaces.Draw + "id") == otherId);
            if (element == null) throw new InvalidDataException("Connector target is missing: " + otherId);
            var otherShape = AttachmentShape(element);
            if (otherShape.IsGroup || otherShape.ElementName is "line" or "connector")
                throw new NotSupportedException("Automatic glue selection requires bounded attachment shapes.");
            OdfRect bounds = otherShape.AttachmentBounds;
            other = otherShape.PageTransform.TransformPoint(new OfficePoint(
                bounds.X.ToPoints() + bounds.Width.ToPoints() / 2, bounds.Y.ToPoints() + bounds.Height.ToPoints() / 2));
        }
        return points.OrderBy(point => DistanceSquared(point, other)).First();
    }

    private bool TryReadSavedEndpoint(bool start, out OfficePoint point) {
        string suffix = start ? "1" : "2";
        string? x = (string?)Element.Attribute(OdfNamespaces.Svg + "x" + suffix), y = (string?)Element.Attribute(OdfNamespaces.Svg + "y" + suffix);
        point = default;
        if (x == null || y == null) return false;
        if (!OdfLength.Parse(x).TryToPoints(out double px) || !OdfLength.Parse(y).TryToPoints(out double py)) return false;
        point = new OfficePoint(px, py); return true;
    }

    private static bool Finite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
    private static double DistanceSquared(OfficePoint a, OfficePoint b) => (a.X - b.X) * (a.X - b.X) + (a.Y - b.Y) * (a.Y - b.Y);
    private static bool Near(OfficePoint a, OfficePoint b) =>
        Math.Abs(a.X - b.X) <= ConnectorEndpointTolerance && Math.Abs(a.Y - b.Y) <= ConnectorEndpointTolerance;
}
