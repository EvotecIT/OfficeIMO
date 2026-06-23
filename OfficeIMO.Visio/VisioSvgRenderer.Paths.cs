using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using OfficeIMO.Drawing;
using Color = OfficeIMO.Drawing.OfficeColor;


namespace OfficeIMO.Visio {
    internal static partial class VisioSvgRenderer {
        private static string BuildPath(VisioPage page, VisioShape shape, IReadOnlyList<(double X, double Y)> localPoints, double scale, bool isClosed) {
            List<OfficePoint> points = new(localPoints.Count);
            for (int i = 0; i < localPoints.Count; i++) {
                (double absX, double absY) = GetPagePoint(shape, localPoints[i].X, localPoints[i].Y);
                (double x, double y) = ToSvg(page, absX, absY, scale);
                points.Add(new OfficePoint(x, y));
            }

            return OfficeSvgFormatting.FormatMoveLinePathData(points, isClosed);
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

        private static double PointsToSvgPixels(double points, double scale) {
            return points * scale / 72D;
        }

        private static string BuildOpenPath(VisioPage page, IReadOnlyList<(double X, double Y)> points, double scale) {
            List<OfficePoint> svgPoints = new(points.Count);
            for (int i = 0; i < points.Count; i++) {
                (double x, double y) = ToSvg(page, points[i].X, points[i].Y, scale);
                svgPoints.Add(new OfficePoint(x, y));
            }

            return OfficeSvgFormatting.FormatMoveLinePathData(svgPoints);
        }

        private static (double X, double Y) ToSvg(VisioPage page, double x, double y, double scale) {
            return (x * scale, (page.Height - y) * scale);
        }

        private static double Distance((double X, double Y) a, (double X, double Y) b) =>
            OfficeIMO.Drawing.OfficeGeometry.Distance(a, b);

        private static double RadiansToDegrees(double radians) => radians * 180D / Math.PI;

        private static string Format(double value) => OfficeSvgFormatting.FormatNumber(value);
    }
}
