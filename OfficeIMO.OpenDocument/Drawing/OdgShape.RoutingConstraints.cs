using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

public sealed partial class OdgShape {
    private sealed class OdgRoutingObstacle {
        internal OdgRoutingObstacle(XElement element, (double Left, double Top, double Right, double Bottom) bounds) { Element = element; Bounds = bounds; }
        internal XElement Element { get; }
        internal (double Left, double Top, double Right, double Bottom) Bounds { get; }
    }

    private (OdgShape? Shape, OfficeConnectorDirections Directions) ReadRoutingAttachment(bool start) {
        string? shapeId = start ? StartShapeId : EndShapeId;
        if (shapeId == null) return (null, OfficeConnectorDirections.Any);
        XElement target = PageElement?.Descendants().FirstOrDefault(element =>
            (string?)element.Attribute(XNamespace.Xml + "id") == shapeId || (string?)element.Attribute(OdfNamespaces.Draw + "id") == shapeId)
            ?? throw new InvalidDataException("Connector target is missing: " + shapeId);
        var shape = new OdgShape((OdgDocument)Document, target);
        string? id = (string?)Element.Attribute(OdfNamespaces.Draw + (start ? "start" : "end") + "-glue-point");
        XElement? point = id == null ? null : target.Elements(OdfNamespaces.Draw + "glue-point")
            .FirstOrDefault(candidate => (string?)candidate.Attribute(OdfNamespaces.Draw + "id") == id);
        OdgGluePointEscapeDirection direction = point == null ? OdgGluePointEscapeDirection.Auto : new OdgGluePoint(shape, point).EscapeDirection;
        return (shape, direction switch {
            OdgGluePointEscapeDirection.Left => OfficeConnectorDirections.Left,
            OdgGluePointEscapeDirection.Right => OfficeConnectorDirections.Right,
            OdgGluePointEscapeDirection.Up => OfficeConnectorDirections.Up,
            OdgGluePointEscapeDirection.Down => OfficeConnectorDirections.Down,
            OdgGluePointEscapeDirection.Horizontal => OfficeConnectorDirections.Horizontal,
            OdgGluePointEscapeDirection.Vertical => OfficeConnectorDirections.Vertical,
            _ => OfficeConnectorDirections.Any
        });
    }

    private static bool RoutingObstacleBlocks(IReadOnlyList<OfficePoint> route, IEnumerable<OdgRoutingObstacle> obstacles,
        XElement? startShape, XElement? endShape) {
        foreach (OdgRoutingObstacle obstacle in obstacles) {
            var box = obstacle.Bounds;
            for (int i = 1; i < route.Count; i++) {
                // Only the terminal segments can pass through their attached shape's inflated bounds.
                // Other segments, including a same-shape loop's middle, must avoid those bounds.
                if ((i == 1 && ReferenceEquals(obstacle.Element, startShape) && InsideRoutingBox(route[0], box)) ||
                    (i == route.Count - 1 && ReferenceEquals(obstacle.Element, endShape) && InsideRoutingBox(route[route.Count - 1], box))) continue;
                if (OfficeGeometry.SegmentIntersectsRectangle((route[i - 1].X, route[i - 1].Y), (route[i].X, route[i].Y),
                    box.Left, box.Top, box.Right, box.Bottom)) return true;
            }
        }
        return false;
    }
    private static bool InsideRoutingBox(OfficePoint point, (double Left, double Top, double Right, double Bottom) box) =>
        point.X >= box.Left && point.X <= box.Right && point.Y >= box.Top && point.Y <= box.Bottom;
}
