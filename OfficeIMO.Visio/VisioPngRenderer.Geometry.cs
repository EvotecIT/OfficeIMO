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
            return VisioConnectorGeometry.GetPoints(connector);
        }


        private static (double X, double Y) GetPagePoint(VisioShape shape, double x, double y) {
            OfficePoint point = VisioNativeShapeTransform.Create(shape).PagePoint(x, y);
            return (point.X, point.Y);
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

        private static (double X, double Y) ToRaster(VisioPage page, double x, double y, VisioRenderProjection projection) =>
            projection.PagePoint(x, y);

        private static (double X, double Y) ToRasterPoint(VisioPage page, VisioShape shape, double x, double y, VisioRenderProjection projection) {
            (double pageX, double pageY) = GetPagePoint(shape, x, y);
            return ToRaster(page, pageX, pageY, projection);
        }

        private static double Distance((double X, double Y) a, (double X, double Y) b) =>
            OfficeGeometry.Distance(a, b);
    }
}
