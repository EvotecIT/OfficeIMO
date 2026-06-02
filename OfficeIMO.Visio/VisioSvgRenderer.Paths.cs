using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Xml;
using Color = OfficeIMO.Drawing.OfficeColor;


namespace OfficeIMO.Visio {
    internal static partial class VisioSvgRenderer {
        private static string BuildPath(VisioPage page, VisioShape shape, IReadOnlyList<(double X, double Y)> localPoints, double scale, bool isClosed) {
            StringBuilder builder = new();
            for (int i = 0; i < localPoints.Count; i++) {
                (double absX, double absY) = GetPagePoint(shape, localPoints[i].X, localPoints[i].Y);
                (double x, double y) = ToSvg(page, absX, absY, scale);
                builder.Append(i == 0 ? "M " : " L ");
                builder.Append(Format(x)).Append(' ').Append(Format(y));
            }

            if (isClosed) {
                builder.Append(" Z");
            }

            return builder.ToString();
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
            double sourceCenterX = (sourceLeft + sourceRight) / 2D;
            double sourceCenterY = (sourceBottom + sourceTop) / 2D;
            double targetCenterX = (targetLeft + targetRight) / 2D;
            double targetCenterY = (targetBottom + targetTop) / 2D;
            double dx = targetCenterX - sourceCenterX;
            double dy = targetCenterY - sourceCenterY;

            if (Math.Abs(dy) > Math.Abs(dx)) {
                x = sourceCenterX;
                y = dy >= 0D ? sourceTop : sourceBottom;
                return;
            }

            x = dx >= 0D ? sourceRight : sourceLeft;
            y = sourceCenterY;
        }

        private static double PointsToSvgPixels(double points, double scale) {
            return points * scale / 72D;
        }

        private static string BuildOpenPath(VisioPage page, IReadOnlyList<(double X, double Y)> points, double scale) {
            StringBuilder builder = new();
            for (int i = 0; i < points.Count; i++) {
                (double x, double y) = ToSvg(page, points[i].X, points[i].Y, scale);
                builder.Append(i == 0 ? "M " : " L ");
                builder.Append(Format(x)).Append(' ').Append(Format(y));
            }

            return builder.ToString();
        }

        private static (double X, double Y) ToSvg(VisioPage page, double x, double y, double scale) {
            return (x * scale, (page.Height - y) * scale);
        }

        private static double Distance((double X, double Y) a, (double X, double Y) b) {
            double dx = b.X - a.X;
            double dy = b.Y - a.Y;
            return Math.Sqrt((dx * dx) + (dy * dy));
        }

        private static double RadiansToDegrees(double radians) => radians * 180D / Math.PI;

        private static string Format(double value) {
            if (Math.Abs(value) < 0.0000001D) value = 0D;
            return value.ToString("0.###", CultureInfo.InvariantCulture);
        }
    }
}
