using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

/// <summary>Initial segment orientation for generated orthogonal Draw routes.</summary>
public enum OdgConnectorRouteOrientation {
    /// <summary>Choose the larger endpoint displacement.</summary>
    Auto,
    /// <summary>Start with a horizontal segment.</summary>
    HorizontalFirst,
    /// <summary>Start with a vertical segment.</summary>
    VerticalFirst
}

public sealed partial class OdgShape {
    /// <summary>
    /// Replaces the route with three orthogonal segments and selects Standard routing.
    /// Offset is in connector-local points. This method does not avoid obstacles or enforce escape directions.
    /// Native applications can subsequently recalculate the route.
    /// </summary>
    public void RouteOrthogonal(OdgConnectorRouteOrientation orientation = OdgConnectorRouteOrientation.Auto, double offset = 0) {
        RequireConnector();
        if (!Enum.IsDefined(typeof(OdgConnectorRouteOrientation), orientation)) throw new ArgumentOutOfRangeException(nameof(orientation));
        var start = new OfficePoint(X1.ToPoints(), Y1.ToPoints()); var end = new OfficePoint(X2.ToPoints(), Y2.ToPoints());
        bool horizontal = orientation == OdgConnectorRouteOrientation.Auto
            ? Math.Abs(end.X - start.X) >= Math.Abs(end.Y - start.Y)
            : orientation == OdgConnectorRouteOrientation.HorizontalFirst;
        SetOrthogonalPoints(OfficeGeometry.CreateOrthogonalConnectorRoute(start, end, horizontal, offset));
    }

    /// <summary>
    /// Searches a bounded set of orthogonal lanes avoiding the declared bounds of unrelated shapes.
    /// Padding is in connector-local points, before transforms. Up to 4096 bounded shapes and 32 lanes per direction are accepted.
    /// Attached shapes and line/connector obstacles are ignored; groups must be supplied as individual bounded children.
    /// No clear lane, unsupported input, or invalid geometry leaves the previous route intact. Native applications can reroute it.
    /// </summary>
    public void RouteOrthogonalAroundShapes(IEnumerable<OdgShape> obstacles, double padding = 6, int maxLanes = 12) =>
        RouteOrthogonalAroundShapes(obstacles, false, padding, maxLanes);

    /// <summary>
    /// Optionally enforces explicit glue-point escape directions while avoiding obstacles and attached shape bounds.
    /// With constraints enabled, routes, padding and departure directions use page coordinates; the saved route is mapped
    /// back into the connector's coordinate space. Free, automatic and standard glue endpoints allow all directions.
    /// Every route must remain orthogonal, retain its departures and match its attachments within 0.1 page point after
    /// native-grid rounding. Unsupported input or an exhausted bounded search leaves the document unchanged.
    /// Native applications may subsequently reroute it.
    /// </summary>
    public void RouteOrthogonalAroundShapes(IEnumerable<OdgShape> obstacles, bool respectEscapeDirections, double padding = 6, int maxLanes = 12) {
        RequireConnector();
        if (obstacles == null) throw new ArgumentNullException(nameof(obstacles));
        if (!Finite(padding) || padding < 0) throw new ArgumentOutOfRangeException(nameof(padding));
        if (maxLanes < 0 || maxLanes > 32) throw new ArgumentOutOfRangeException(nameof(maxLanes), "Lane count must be between 0 and 32.");
        OdgShape[] shapes = obstacles.Take(4097).ToArray();
        if (shapes.Length > 4096) throw new ArgumentException("At most 4096 obstacle shapes are accepted.", nameof(obstacles));
        OfficeTransform inverse = InversePageTransform();
        OfficeTransform pageTransform = PageTransform;
        OfficeTransform routingTransform = respectEscapeDirections ? OfficeTransform.Identity : inverse;
        var startAttachment = respectEscapeDirections ? ReadRoutingAttachment(true) : default;
        var endAttachment = respectEscapeDirections ? ReadRoutingAttachment(false) : default;
        var bounds = new List<OdgRoutingObstacle>();
        var seen = new HashSet<XElement>();
        foreach (OdgShape shape in shapes.Concat(new[] { startAttachment.Shape, endAttachment.Shape }.OfType<OdgShape>())) {
            if (shape == null) throw new ArgumentException("Obstacle shapes cannot contain null entries.", nameof(obstacles));
            if (!ReferenceEquals(Document, shape.Document) || PageElement == null || !ReferenceEquals(PageElement, shape.PageElement))
                throw new ArgumentException("Obstacles must belong to the connector's drawing page.", nameof(obstacles));
            if (!seen.Add(shape.Element) || shape.ElementName is "line" or "connector" ||
                (!respectEscapeDirections && shape.XmlId != null && (shape.XmlId == StartShapeId || shape.XmlId == EndShapeId))) continue;
            if (shape.IsGroup || shape.ElementName is not ("rect" or "ellipse" or "circle" or "frame" or "custom-shape" or "path" or "polygon" or "polyline"))
                throw new NotSupportedException("Obstacle routing requires individual bounded shapes; expand groups into their bounded children.");
            OdfRect box = shape.Bounds;
            double x = box.X.ToPoints(), y = box.Y.ToPoints(), w = box.Width.ToPoints(), h = box.Height.ToPoints();
            if (!Finite(x) || !Finite(y) || !Finite(w) || !Finite(h) || w < 0 || h < 0)
                throw new InvalidDataException("Obstacle bounds must be finite and nonnegative.");
            OfficeTransform transform = shape.PageTransform.Then(routingTransform);
            OfficePoint[] corners = new[] { new OfficePoint(x, y), new OfficePoint(x + w, y), new OfficePoint(x, y + h), new OfficePoint(x + w, y + h) }
                .Select(transform.TransformPoint).ToArray();
            // Reserve endpoint/grid tolerance so native coordinate rounding cannot cross an obstacle boundary.
            double clearance = padding + ConnectorEndpointTolerance;
            var inflated = (Left: corners.Min(point => point.X) - clearance, Top: corners.Min(point => point.Y) - clearance,
                Right: corners.Max(point => point.X) + clearance, Bottom: corners.Max(point => point.Y) + clearance);
            if (!Finite(inflated.Left) || !Finite(inflated.Top) || !Finite(inflated.Right) || !Finite(inflated.Bottom))
                throw new InvalidDataException("Transformed obstacle bounds exceed finite coordinates.");
            bounds.Add(new OdgRoutingObstacle(shape.Element, inflated));
        }
        var start = new OfficePoint(X1.ToPoints(), Y1.ToPoints()); var end = new OfficePoint(X2.ToPoints(), Y2.ToPoints());
        if (respectEscapeDirections) { start = pageTransform.TransformPoint(start); end = pageTransform.TransformPoint(end); }
        var startBounds = bounds.FirstOrDefault(box => ReferenceEquals(box.Element, startAttachment.Shape?.Element))?.Bounds;
        var endBounds = bounds.FirstOrDefault(box => ReferenceEquals(box.Element, endAttachment.Shape?.Element))?.Bounds;
        double exitClearance = ConnectorEndpointTolerance + 36D / 2540 * Math.Max(
            Math.Abs(pageTransform.M11) + Math.Abs(pageTransform.M21), Math.Abs(pageTransform.M12) + Math.Abs(pageTransform.M22));
        IEnumerable<OfficePoint[]> candidates = respectEscapeDirections
            ? OfficeGeometry.EnumerateConstrainedOrthogonalConnectorRoutes(start, end, startAttachment.Directions, endAttachment.Directions,
                Math.Max(padding, 6), maxLanes, exitClearance, startBounds, endBounds)
            : OfficeGeometry.EnumerateOrthogonalConnectorRoutes(start, end, Math.Max(padding, 6), maxLanes);
        foreach (OfficePoint[] points in candidates) {
            OfficePoint[] local = respectEscapeDirections ? points.Select(inverse.TransformPoint).ToArray() : points;
            OfficePoint[] resolved = respectEscapeDirections ? local.Select(point => {
                OfficePoint native = NativeConnectorPoint(point);
                return pageTransform.TransformPoint(new OfficePoint(native.X * 72 / 2540, native.Y * 72 / 2540));
            }).ToArray() : points;
            if (respectEscapeDirections && (!Near(resolved[0], start) || !Near(resolved[resolved.Length - 1], end) ||
                !OfficeGeometry.MatchesConnectorDirections(resolved, startAttachment.Directions, endAttachment.Directions, ConnectorEndpointTolerance))) continue;
            if (!RoutingObstacleBlocks(resolved, bounds, startAttachment.Shape?.Element, endAttachment.Shape?.Element)) {
                SetOrthogonalPoints(local); return;
            }
        }
        throw new NotSupportedException("No clear route satisfying the requested constraints was found within the lane search. Increase the lane count, adjust padding or endpoints, or assign an explicit route.");
    }

    private void SetOrthogonalPoints(IReadOnlyList<OfficePoint> points) => SetConnectorRouteCore(
        points.Select((point, index) => index == 0 ? OfficePathCommand.MoveTo(point) : OfficePathCommand.LineTo(point)), OdgConnectorKind.Standard);
}
