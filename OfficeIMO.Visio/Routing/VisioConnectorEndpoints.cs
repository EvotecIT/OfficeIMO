using System;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio;

internal sealed class VisioConnectorAttachment {
    internal VisioConnectorAttachment(VisioShape shape, OfficePoint point) {
        LocalPoint = VisioConnectorEndpoints.ToLocalPoint(shape, point);
        Width = shape.WidthInches; Height = shape.HeightInches;
    }
    internal OfficePoint LocalPoint { get; }
    internal double Width { get; }
    internal double Height { get; }
}

/// <summary>One page-coordinate endpoint resolver for editing, rendering, routing and serialization.</summary>
internal static class VisioConnectorEndpoints {
    internal static OfficePoint Resolve(VisioConnector connector, bool start) {
        VisioShape? shape = start ? connector.From : connector.To;
        if (shape == null) return connector.GetFreePoint(start);
        VisioConnectionPoint? point = start ? connector.FromConnectionPoint : connector.ToConnectionPoint;
        if (point != null) return ToPagePoint(shape, new OfficePoint(point.X, point.Y));
        VisioConnectorAttachment? anchor = start ? connector.StartAttachment : connector.EndAttachment;
        if (anchor != null) return ToPagePoint(shape, new OfficePoint(
            anchor.LocalPoint.X * (Math.Abs(anchor.Width) > 1e-12 ? shape.WidthInches / anchor.Width : 1),
            anchor.LocalPoint.Y * (Math.Abs(anchor.Height) > 1e-12 ? shape.HeightInches / anchor.Height : 1)));
        VisioShapeBounds bounds = shape.GetPageShapeBounds();
        VisioShape? other = start ? connector.To : connector.From;
        VisioShapeBounds target = other != null ? other.GetPageShapeBounds() : PointBounds(connector.GetFreePoint(!start));
        return OfficeGeometry.ResolveRectangleBoundaryEndpoint(bounds.LeftInches, bounds.BottomInches, bounds.RightInches, bounds.TopInches,
            target.LeftInches, target.BottomInches, target.RightInches, target.TopInches);
    }

    internal static void Resolve(VisioConnector connector, out double startX, out double startY, out double endX, out double endY) {
        OfficePoint start = connector.StartPoint, end = connector.EndPoint;
        startX = start.X; startY = start.Y; endX = end.X; endY = end.Y;
    }

    internal static VisioShapeBounds Bounds(VisioConnector connector, bool start) =>
        (start ? connector.From : connector.To)?.GetPageShapeBounds() ?? PointBounds(start ? connector.StartPoint : connector.EndPoint);

    private static VisioShapeBounds PointBounds(OfficePoint point) => new(point.X, point.Y, point.X, point.Y);

    internal static OfficePoint ToPagePoint(VisioShape shape, OfficePoint point) =>
        VisioNativeShapeTransform.Create(shape).PagePoint(point.X, point.Y);

    internal static OfficePoint ToLocalPoint(VisioShape shape, OfficePoint point) =>
        VisioNativeShapeTransform.Create(shape).LocalPoint(point);
}
