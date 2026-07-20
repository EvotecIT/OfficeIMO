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

        private static void DrawStencilArtwork(RasterCanvas canvas, VisioPage page, VisioShape shape) {
            string? stencilKey = VisioStencilArtwork.GetKey(shape);
            if (string.IsNullOrEmpty(stencilKey)) {
                return;
            }

            double placementScale = string.IsNullOrWhiteSpace(shape.Text) ? 0.58D : 0.34D;
            double iconSize = Math.Max(0.08D, Math.Min(shape.Width, shape.Height) * placementScale);
            double localCx = shape.Width / 2D;
            double localCy = string.IsNullOrWhiteSpace(shape.Text)
                ? shape.Height / 2D
                : shape.Height - Math.Min(shape.Height * 0.28D, iconSize * 0.72D);
            (double cx, double cy) = GetPagePoint(shape, localCx, localCy);
            (double x, double y) = ToRaster(page, cx, cy, canvas.Scale);
            double size = iconSize * canvas.Scale;
            Color color = VisioStencilArtwork.ResolveColor(shape, 155);
            double stroke = Math.Max(canvas.Supersampling, size * 0.045D);
            double rasterRotation = ToRasterRotation(shape.Angle);
            (double X, double Y) Point(double offsetX, double offsetY) =>
                OfficeGeometry.RotatePoint((x + (size * offsetX), y + (size * offsetY)), x, y, -rasterRotation);
            (double X, double Y)[] Points(params (double X, double Y)[] offsets) {
                (double X, double Y)[] points = new (double X, double Y)[offsets.Length];
                for (int i = 0; i < offsets.Length; i++) {
                    points[i] = Point(offsets[i].X, offsets[i].Y);
                }

                return points;
            }

            switch (stencilKey) {
                case "person":
                    (double headX, double headY) = Point(0D, -0.18D);
                    canvas.DrawEllipse(headX, headY, size * 0.16D, size * 0.16D, Color.Transparent, color, stroke);
                    StrokeArc(canvas, x, y + size * 0.22D, size * 0.31D, size * 0.24D, 205D, 335D, color, stroke, rasterRotation, x, y);
                    break;
                case "data":
                    StrokeEllipse(canvas, x, y - size * 0.18D, size * 0.31D, size * 0.11D, color, stroke, rasterRotation, x, y);
                    StrokePolyline(canvas, Points((-0.31D, -0.18D), (-0.31D, 0.26D)), color, stroke);
                    StrokePolyline(canvas, Points((0.31D, -0.18D), (0.31D, 0.26D)), color, stroke);
                    StrokeArc(canvas, x, y + size * 0.26D, size * 0.31D, size * 0.11D, 0D, 180D, color, stroke, rasterRotation, x, y);
                    break;
                case "security":
                    StrokePolyline(canvas, Points(
                        (0D, -0.36D),
                        (0.3D, -0.22D),
                        (0.22D, 0.22D),
                        (0D, 0.38D),
                        (-0.22D, 0.22D),
                        (-0.3D, -0.22D),
                        (0D, -0.36D)), color, stroke);
                    break;
                case "compute":
                    StrokePolyline(canvas, Points(
                        (-0.34D, -0.24D),
                        (0.34D, -0.24D),
                        (0.34D, 0.24D),
                        (-0.34D, 0.24D),
                        (-0.34D, -0.24D)), color, stroke);
                    StrokePolyline(canvas, Points((-0.22D, -0.06D), (0.22D, -0.06D)), color, stroke);
                    StrokePolyline(canvas, Points((-0.22D, 0.08D), (0.22D, 0.08D)), color, stroke);
                    break;
                case "cloud":
                    StrokeEllipse(canvas, x - size * 0.16D, y + size * 0.02D, size * 0.2D, size * 0.15D, color, stroke, rasterRotation, x, y);
                    StrokeEllipse(canvas, x + size * 0.08D, y - size * 0.06D, size * 0.24D, size * 0.2D, color, stroke, rasterRotation, x, y);
                    StrokeEllipse(canvas, x + size * 0.25D, y + size * 0.05D, size * 0.16D, size * 0.12D, color, stroke, rasterRotation, x, y);
                    StrokePolyline(canvas, Points((-0.33D, 0.16D), (0.37D, 0.16D)), color, stroke);
                    break;
                case "container":
                    StrokePolyline(canvas, Points(
                        (0D, -0.36D),
                        (0.3096D, -0.18D),
                        (0.3096D, 0.18D),
                        (0D, 0.36D),
                        (-0.3096D, 0.18D),
                        (-0.3096D, -0.18D),
                        (0D, -0.36D)), color, stroke);
                    break;
                case "event":
                    StrokePolyline(canvas, Points((-0.32D, -0.16D), (0.28D, -0.16D)), color, stroke);
                    StrokePolyline(canvas, Points((-0.32D, 0D), (0.18D, 0D)), color, stroke);
                    StrokePolyline(canvas, Points((-0.32D, 0.16D), (0.28D, 0.16D)), color, stroke);
                    break;
                case "monitoring":
                    StrokePolyline(canvas, Points(
                        (-0.36D, 0D),
                        (-0.14D, 0D),
                        (-0.04D, -0.22D),
                        (0.09D, 0.2D),
                        (0.19D, 0D),
                        (0.36D, 0D)), color, stroke);
                    break;
            }
        }

        private static bool DrawPackagePreviewArtwork(RasterCanvas canvas, VisioPage page, VisioShape shape, VisioPngSaveOptions options) {
            if (!VisioPackagePreviewArtwork.TryGetRasterImage(
                    shape,
                    options.ImageCodec,
                    canvas.OutlineFont,
                    canvas.Fonts,
                    canvas.TextShapingProvider,
                    canvas.TextShapingLanguage,
                    options.ImageDiagnostics,
                    options.ImageDiagnosticSource,
                    canvas.CancellationToken,
                    out OfficeRasterImage? raster) ||
                raster == null) {
                return false;
            }

            double placementScale = string.IsNullOrWhiteSpace(shape.Text) ? 0.64D : 0.42D;
            double imageWidth = Math.Max(0.01D, shape.Width * placementScale);
            double imageHeight = Math.Max(0.01D, shape.Height * placementScale);
            double localCx = shape.Width / 2D;
            double localCy = string.IsNullOrWhiteSpace(shape.Text)
                ? shape.Height / 2D
                : shape.Height - Math.Min(shape.Height * 0.3D, imageHeight * 0.72D);
            (double cx, double cy) = GetPagePoint(shape, localCx, localCy);
            (double centerX, double centerY) = ToRaster(page, cx, cy, canvas.Scale);
            double targetWidth = imageWidth * canvas.Scale;
            double targetHeight = imageHeight * canvas.Scale;
            OfficeImageRenderPlan renderPlan = OfficeImageRenderPlan.CreateTopLeft(
                raster.Width,
                raster.Height,
                centerX - (targetWidth / 2D),
                centerY - (targetHeight / 2D),
                targetWidth,
                targetHeight,
                OfficeImageFit.Contain);

            canvas.DrawImage(
                raster,
                renderPlan.ToVisibleProjection(
                    rotationDegrees: -OfficeGeometry.RadiansToDegrees(ToRasterRotation(shape.Angle)),
                    rotationCenterX: centerX,
                    rotationCenterY: centerY));
            return true;
        }

        private static double ToRasterRotation(double visioRadians) => -visioRadians;

        private static Color ApplyBackgroundTransparency(Color color, double? transparency) {
            if (!transparency.HasValue) {
                return color;
            }

            double clamped = Math.Max(0D, Math.Min(100D, transparency.Value));
            byte alpha = (byte)Math.Round(color.A * (1D - (clamped / 100D)));
            return Color.FromRgba(color.R, color.G, color.B, alpha);
        }
    }
}
