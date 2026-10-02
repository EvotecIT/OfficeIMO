using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio {
    internal static partial class VisioSvgPreviewRasterizer {
        private static bool RenderRectangle(OfficeRasterCanvas canvas, XElement element, SvgPaint paint, SvgTransform transform, SvgRenderContext context) {
            double x = ReadLength(element, "x", 0D, context, SvgLengthAxis.X);
            double y = ReadLength(element, "y", 0D, context, SvgLengthAxis.Y);
            double width = ReadLength(element, "width", 0D, context, SvgLengthAxis.X);
            double height = ReadLength(element, "height", 0D, context, SvgLengthAxis.Y);
            if (width <= 0D || height <= 0D) {
                return false;
            }

            bool hasRx = TryParseLength(element.Attribute("rx")?.Value, GetLengthReference(context, SvgLengthAxis.X), out double rx);
            bool hasRy = TryParseLength(element.Attribute("ry")?.Value, GetLengthReference(context, SvgLengthAxis.Y), out double ry);
            if (hasRx && !hasRy) {
                ry = rx;
            } else if (!hasRx && hasRy) {
                rx = ry;
            }

            rx = Math.Min(Math.Abs(rx), width / 2D);
            ry = Math.Min(Math.Abs(ry), height / 2D);
            List<(double X, double Y)> points = rx > 0D && ry > 0D
                ? CreateRoundedRectanglePoints(x, y, width, height, rx, ry, transform.CurveScale)
                : new List<(double X, double Y)> {
                    (x, y),
                    (x + width, y),
                    (x + width, y + height),
                    (x, y + height)
                };

            return RenderPolyline(canvas, points, closed: true, paint, transform);
        }

        private static bool RenderEllipse(OfficeRasterCanvas canvas, double cx, double cy, double rx, double ry, SvgPaint paint, SvgTransform transform) {
            if (rx <= 0D || ry <= 0D) {
                return false;
            }

            IReadOnlyList<OfficePoint> ellipsePoints = CreateEllipsePoints(cx, cy, rx, ry, transform);
            if (paint.HasFill) {
                if (paint.FillRadialGradient != null || paint.FillGradient != null) {
                    if (paint.FillRadialGradient != null) {
                        canvas.FillRadialGradientPolygon(ellipsePoints, paint.FillRadialGradient);
                    } else {
                        canvas.FillLinearGradientPolygon(ellipsePoints, paint.FillGradient!);
                    }
                } else {
                    canvas.FillPolygon(ellipsePoints, paint.Fill);
                }
            }

            if (paint.HasStroke && paint.StrokeWidth > 0D) {
                double strokeScale = GetStrokeScale(paint, transform);
                double strokeWidth = paint.StrokeWidth * strokeScale;
                IReadOnlyList<double>? dashPattern = ScaleDashPattern(paint.DashPattern, strokeScale);
                StrokeClosedContour(canvas, ellipsePoints, paint, strokeWidth, dashPattern, paint.DashOffset * strokeScale, transform);
            }

            return paint.HasFill || paint.HasStroke;
        }

        private static bool RenderPath(OfficeRasterCanvas canvas, XElement element, IReadOnlyList<SvgPathContour> contours, SvgPaint paint, SvgTransform transform, SvgRenderContext context) {
            if (contours.Count == 0) {
                return false;
            }

            bool rendered = false;
            List<IReadOnlyList<OfficePoint>> closedContours = new();
            List<(IReadOnlyList<OfficePoint> Points, bool Closed)> projectedContours = new(contours.Count);
            for (int i = 0; i < contours.Count; i++) {
                List<OfficePoint> projected = ProjectPoints(contours[i].Points, transform);
                if (projected.Count < 2) {
                    continue;
                }

                bool closed = contours[i].IsClosed && projected.Count >= 3;
                projectedContours.Add((projected, closed));
                if (projected.Count >= 3) {
                    closedContours.Add(projected);
                }
            }

            if (paint.HasFill && closedContours.Count > 0) {
                bool useEvenOddFill = context.CurrentFillRule == OfficeFillRule.EvenOdd;
                if (paint.FillRadialGradient != null) {
                    FillGradientContours(canvas, closedContours, null, paint.FillRadialGradient, useEvenOddFill);
                } else if (paint.FillGradient != null) {
                    FillGradientContours(canvas, closedContours, paint.FillGradient, null, useEvenOddFill);
                } else if (useEvenOddFill) {
                    if (closedContours.Count > 1) {
                        canvas.FillPolygonsEvenOdd(closedContours, paint.Fill);
                    } else {
                        canvas.FillPolygon(closedContours[0], paint.Fill);
                    }
                } else {
                    canvas.FillPolygonsNonZero(closedContours, paint.Fill);
                }

                rendered = true;
            }

            if (paint.HasStroke && paint.StrokeWidth > 0D) {
                double strokeScale = GetStrokeScale(paint, transform);
                double strokeWidth = paint.StrokeWidth * strokeScale;
                IReadOnlyList<double>? dashPattern = ScaleDashPattern(paint.DashPattern, strokeScale);
                var strokes = new List<OfficeFlattenedPathContour>(projectedContours.Count);
                foreach (var contour in projectedContours) strokes.Add(new OfficeFlattenedPathContour(contour.Points, contour.Closed));
                StrokePreviewContours(canvas, strokes, paint, strokeWidth, dashPattern, paint.DashOffset * strokeScale, transform);

                rendered = true;
            }

            return rendered;
        }

        private static void FillGradientContours(
            OfficeRasterCanvas canvas,
            IReadOnlyList<IReadOnlyList<OfficePoint>> contours,
            OfficeLinearGradient? linearGradient,
            OfficeRadialGradient? radialGradient,
            bool useEvenOddFill) {
            canvas.FillContourPaint(contours, useEvenOddFill ? OfficeFillRule.EvenOdd : OfficeFillRule.NonZero,
                canvas.CreateContourPaint(contours, OfficeColor.Transparent, linearGradient, radialGradient));
        }
        private static bool RenderPolyline(OfficeRasterCanvas canvas, IReadOnlyList<(double X, double Y)> points, bool closed, SvgPaint paint, SvgTransform transform) {
            if (points.Count < 2) {
                return false;
            }

            List<OfficePoint> projected = ProjectPoints(points, transform);

            bool filled = projected.Count >= 3 && paint.HasFill;
            if (filled) {
                if (paint.FillRadialGradient != null) {
                    canvas.FillRadialGradientPolygon(projected, paint.FillRadialGradient);
                } else if (paint.FillGradient != null) {
                    canvas.FillLinearGradientPolygon(projected, paint.FillGradient);
                } else {
                    canvas.FillPolygon(projected, paint.Fill);
                }
            }

            if (paint.HasStroke && paint.StrokeWidth > 0D) {
                double strokeScale = GetStrokeScale(paint, transform);
                double strokeWidth = paint.StrokeWidth * strokeScale;
                IReadOnlyList<double>? dashPattern = ScaleDashPattern(paint.DashPattern, strokeScale);
                if (closed && projected.Count >= 3) {
                    StrokeClosedContour(canvas, projected, paint, strokeWidth, dashPattern, paint.DashOffset * strokeScale, transform);
                } else {
                    StrokeOpenContour(canvas, projected, paint, strokeWidth, dashPattern, paint.DashOffset * strokeScale, transform);
                }
            }

            return filled || paint.HasStroke;
        }

        private static IReadOnlyList<double>? ScaleDashPattern(IReadOnlyList<double>? pattern, double scale) {
            if (pattern == null || pattern.Count == 0) {
                return null;
            }

            List<double> scaled = new(pattern.Count);
            for (int i = 0; i < pattern.Count; i++) {
                double value = pattern[i] * Math.Max(0.0001D, scale);
                if (value >= 0D && !double.IsNaN(value) && !double.IsInfinity(value)) {
                    scaled.Add(value);
                }
            }

            return scaled.Count == 0 ? null : scaled;
        }

        private static double GetStrokeScale(SvgPaint paint, SvgTransform transform) =>
            paint.NonScalingStroke ? 1D : transform.StrokeScale;

        private static IReadOnlyList<OfficePoint> CreateEllipsePoints(double cx, double cy, double rx, double ry, SvgTransform transform) {
            var points = OfficeCurveFlattening.Ellipse(cx, cy, rx, ry, transform.CurveScale);
            for (int i = 0; i < points.Count; i++) points[i] = transform.Apply(points[i].X, points[i].Y);
            return points;
        }

        private static List<(double X, double Y)> CreateRoundedRectanglePoints(double x, double y, double width, double height, double rx, double ry, double pixelsPerUnit) {
            var points = new List<(double X, double Y)>();
            foreach (OfficePoint point in OfficeCurveFlattening.RoundedRectangle(x, y, width, height, rx, ry, pixelsPerUnit)) points.Add((point.X, point.Y));
            return points;
        }
        private static List<OfficePoint> ProjectPoints(IReadOnlyList<(double X, double Y)> points, SvgTransform transform) {
            List<OfficePoint> projected = new(points.Count);
            for (int i = 0; i < points.Count; i++) {
                projected.Add(transform.Apply(points[i].X, points[i].Y));
            }

            return projected;
        }

    }
}
