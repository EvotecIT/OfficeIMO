using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using OfficeIMO.Drawing;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Visio {
    internal static partial class VisioPngRenderer {

        private static void StrokeLine(RasterCanvas canvas, double x1, double y1, double x2, double y2, Color color, double width) =>
            canvas.StrokePolyline(new[] { (x1, y1), (x2, y2) }, color, width, OfficeStrokeDashStyle.Solid);

        private static void StrokeRect(RasterCanvas canvas, double x, double y, double width, double height, Color color, double stroke) =>
            StrokePolyline(canvas, new[] { (x, y), (x + width, y), (x + width, y + height), (x, y + height), (x, y) }, color, stroke);

        private static void StrokeEllipse(RasterCanvas canvas, double x, double y, double rx, double ry, Color color, double stroke, double rotationRadians = 0D, double rotationCenterX = 0D, double rotationCenterY = 0D) =>
            canvas.DrawEllipse(x, y, rx, ry, Color.Transparent, color, stroke, OfficeStrokeDashStyle.Solid, rotationRadians, rotationCenterX, rotationCenterY);

        private static void StrokeArc(RasterCanvas canvas, double x, double y, double rx, double ry, double startDegrees, double endDegrees, Color color, double stroke, double rotationRadians = 0D, double rotationCenterX = 0D, double rotationCenterY = 0D) =>
            canvas.DrawArc(x, y, rx, ry, startDegrees, endDegrees, color, stroke, rotationRadians, rotationCenterX, rotationCenterY);

        private static void StrokePolyline(RasterCanvas canvas, IReadOnlyList<(double X, double Y)> points, Color color, double stroke) =>
            canvas.StrokePolyline(points, color, stroke, OfficeStrokeDashStyle.Solid);

        private static IReadOnlyList<(double X, double Y)> GetHexPoints(double x, double y, double size) {
            double r = size * 0.36D;
            return new[] {
                (x, y - r),
                (x + r * 0.86D, y - r * 0.5D),
                (x + r * 0.86D, y + r * 0.5D),
                (x, y + r),
                (x - r * 0.86D, y + r * 0.5D),
                (x - r * 0.86D, y - r * 0.5D),
                (x, y - r)
            };
        }

        private static List<(double X, double Y)> GetConnectorPoints(VisioConnector connector) {
            ComputeConnectorEndpoints(connector, out double startX, out double startY, out double endX, out double endY);
            List<(double X, double Y)> waypoints = new(connector.Waypoints.Count);
            if (connector.Waypoints.Count > 0) {
                foreach (VisioConnectorWaypoint waypoint in connector.Waypoints) {
                    waypoints.Add((waypoint.X, waypoint.Y));
                }
            }

            return OfficeGeometry.BuildConnectorPolyline(
                (startX, startY),
                (endX, endY),
                waypoints,
                connector.Kind == ConnectorKind.RightAngle);
        }

        private static void ComputeConnectorEndpoints(VisioConnector connector, out double startX, out double startY, out double endX, out double endY) {
            if (connector.FromConnectionPoint != null) {
                (startX, startY) = GetPagePoint(connector.From, connector.FromConnectionPoint.X, connector.FromConnectionPoint.Y);
            } else {
                (double fromLeft, double fromBottom, double fromRight, double fromTop) = GetPageBounds(connector.From);
                (double toLeft, double toBottom, double toRight, double toTop) = GetPageBounds(connector.To);
                ResolveFallbackEndpoint(fromLeft, fromBottom, fromRight, fromTop, toLeft, toBottom, toRight, toTop, out startX, out startY);
            }

            if (connector.ToConnectionPoint != null) {
                (endX, endY) = GetPagePoint(connector.To, connector.ToConnectionPoint.X, connector.ToConnectionPoint.Y);
            } else {
                (double toLeft, double toBottom, double toRight, double toTop) = GetPageBounds(connector.To);
                (double fromLeft, double fromBottom, double fromRight, double fromTop) = GetPageBounds(connector.From);
                ResolveFallbackEndpoint(toLeft, toBottom, toRight, toTop, fromLeft, fromBottom, fromRight, fromTop, out endX, out endY);
            }
        }

        private static (double X, double Y) ResolveConnectorLabelPoint(VisioConnector connector, IReadOnlyList<(double X, double Y)> points) {
            VisioConnectorLabelPlacement? placement = connector.LabelPlacement;
            if (placement?.AbsolutePinX.HasValue == true && placement.AbsolutePinY.HasValue) {
                return (placement.AbsolutePinX.Value, placement.AbsolutePinY.Value);
            }

            double position = VisioConnectorLabelPlacement.ClampPosition(placement?.Position ?? 0.5D);
            (double x, double y) = OfficeGeometry.InterpolatePolyline(points, position);
            return (x + (placement?.OffsetX ?? 0D), y + (placement?.OffsetY ?? 0D));
        }

        private static VisioRenderConnectorLabelPlacement ResolveConnectorLabel(VisioConnector connector, IReadOnlyList<(double X, double Y)> points) {
            (double x, double y) = ResolveConnectorLabelPoint(connector, points);
            VisioConnectorLabelPlacement? placement = connector.LabelPlacement;
            double width = Math.Max(0.6D, connector.TextStyle?.TextWidth ?? placement?.Width ?? 1.35D);
            double height = Math.Max(0.18D, connector.TextStyle?.TextHeight ?? placement?.Height ?? 0.34D);
            return new VisioRenderConnectorLabelPlacement(x, y, width, height, adjusted: false);
        }

        private static (double X, double Y) GetPagePoint(VisioShape shape, double x, double y) {
            (double absX, double absY) = shape.GetAbsolutePoint(x, y);
            return shape.Parent != null
                ? GetPagePoint(shape.Parent, absX, absY)
                : (absX, absY);
        }

        private static (double Left, double Bottom, double Right, double Top) GetPageBounds(VisioShape shape) {
            (double x1, double y1) = GetPagePoint(shape, 0, 0);
            (double x2, double y2) = GetPagePoint(shape, shape.Width, 0);
            (double x3, double y3) = GetPagePoint(shape, 0, shape.Height);
            (double x4, double y4) = GetPagePoint(shape, shape.Width, shape.Height);
            double left = Math.Min(Math.Min(x1, x2), Math.Min(x3, x4));
            double right = Math.Max(Math.Max(x1, x2), Math.Max(x3, x4));
            double bottom = Math.Min(Math.Min(y1, y2), Math.Min(y3, y4));
            double top = Math.Max(Math.Max(y1, y2), Math.Max(y3, y4));
            return (left, bottom, right, top);
        }

        private static void ResolveFallbackEndpoint(
            double sourceLeft,
            double sourceBottom,
            double sourceRight,
            double sourceTop,
            double targetLeft,
            double targetBottom,
            double targetRight,
            double targetTop,
            out double x,
            out double y) {
            OfficeGeometry.ResolveRectangleBoundaryEndpoint(
                sourceLeft,
                sourceBottom,
                sourceRight,
                sourceTop,
                targetLeft,
                targetBottom,
                targetRight,
                targetTop,
                out x,
                out y);
        }

        private static (double X, double Y) ToRaster(VisioPage page, double x, double y, double scale) =>
            (x * scale, (page.Height - y) * scale);

        private static (double X, double Y) ToRasterPoint(VisioPage page, VisioShape shape, double x, double y, double scale) {
            (double pageX, double pageY) = GetPagePoint(shape, x, y);
            return ToRaster(page, pageX, pageY, scale);
        }

        private static double Distance((double X, double Y) a, (double X, double Y) b) =>
            OfficeGeometry.Distance(a, b);
    }
}
