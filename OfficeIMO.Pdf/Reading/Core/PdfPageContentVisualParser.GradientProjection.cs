using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfPageContentVisualParser {
    private sealed partial class Parser {
        internal static void CreateShadingGradients(
            PdfPageShadingResource shading,
            double x, double y, double width, double height,
            Matrix2D transform, double pageHeight,
            out OfficeLinearGradient? linearGradient,
            out OfficeRadialGradient? radialGradient) {
            linearGradient = null;
            radialGradient = null;
            double paintWidth = Math.Max(width, 0.0001D);
            double paintHeight = Math.Max(height, 0.0001D);
            if (!shading.IsRadial) {
                if (shading.X0 == shading.X1 && shading.Y0 == shading.Y1) return;
                OfficeTransform coordinates = new OfficeTransform(transform.A, transform.B,
                    transform.C, transform.D, transform.E, transform.F)
                    .Then(new OfficeTransform(1D / paintWidth, 0D, 0D, -1D / paintHeight,
                        -x / paintWidth, (pageHeight - y) / paintHeight));
                // Preserve the color field, including portions whose endpoints
                // lie outside the clipping rectangle. Clipping the axis itself
                // changes diagonal colors even when stops are resampled.
                try {
                    linearGradient = OfficeLinearGradient.CreateImported(shading.X0, shading.Y0,
                        shading.X1, shading.Y1, shading.Stops).TransformCoordinates(coordinates);
                } catch (ArgumentException) {
                    // The calling parser reports an unrepresentable paint rather
                    // than inventing a horizontal gradient for degenerate input.
                    linearGradient = null;
                }
                return;
            }

            (double X, double Y) start = transform.Transform(shading.X0, shading.Y0);
            (double X, double Y) end = transform.Transform(shading.X1, shading.Y1);
            double startX = (start.X - x) / paintWidth;
            double startY = ((pageHeight - start.Y) - y) / paintHeight;
            double endX = (end.X - x) / paintWidth;
            double endY = ((pageHeight - end.Y) - y) / paintHeight;
            double startRadiusX = TransformRadiusX(transform, shading.R0) / paintWidth;
            double startRadiusY = TransformRadiusY(transform, shading.R0) / paintHeight;
            double endRadiusX = TransformRadiusX(transform, shading.R1) / paintWidth;
            double endRadiusY = TransformRadiusY(transform, shading.R1) / paintHeight;
            if (NearlyEqual(startX, endX) && NearlyEqual(startY, endY)
                && NearlyEqual(startRadiusX, endRadiusX) && NearlyEqual(startRadiusY, endRadiusY)) {
                endRadiusX = startRadiusX + 0.5D;
                endRadiusY = startRadiusY + 0.5D;
            }
            radialGradient = endRadiusX > 0D && endRadiusY > 0D
                ? new OfficeRadialGradient(startX, startY, startRadiusX, startRadiusY,
                    endX, endY, endRadiusX, endRadiusY, shading.Stops)
                : new OfficeRadialGradient(startX, startY, Math.Max(startRadiusX, startRadiusY),
                    endX, endY, Math.Max(endRadiusX, endRadiusY), shading.Stops);
        }

        private static double TransformRadiusX(Matrix2D transform, double radius) =>
            TransformRadius(radius, Math.Sqrt((transform.A * transform.A) + (transform.B * transform.B)));

        private static double TransformRadiusY(Matrix2D transform, double radius) =>
            TransformRadius(radius, Math.Sqrt((transform.C * transform.C) + (transform.D * transform.D)));

        private static double TransformRadius(double radius, double scale) {
            if (radius <= 0D) return 0D;
            return !double.IsNaN(scale) && !double.IsInfinity(scale) && scale > 0D ? radius * scale : radius;
        }
    }
}
