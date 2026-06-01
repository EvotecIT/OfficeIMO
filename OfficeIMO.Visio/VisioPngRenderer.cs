using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using OfficeIMO.Drawing;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Visio {
    internal static class VisioPngRenderer {
        private static readonly byte[] PngSignature = { 137, 80, 78, 71, 13, 10, 26, 10 };

        public static byte[] Render(VisioPage page, VisioPngSaveOptions options) {
            if (options.PixelsPerInch <= 0D || double.IsNaN(options.PixelsPerInch) || double.IsInfinity(options.PixelsPerInch)) {
                throw new ArgumentOutOfRangeException(nameof(options), "PixelsPerInch must be a finite positive number.");
            }

            if (options.Supersampling < 1 || options.Supersampling > 4) {
                throw new ArgumentOutOfRangeException(nameof(options), "Supersampling must be between 1 and 4.");
            }

            int width = Math.Max(1, (int)Math.Ceiling(Math.Max(page.Width, 0.01D) * options.PixelsPerInch));
            int height = Math.Max(1, (int)Math.Ceiling(Math.Max(page.Height, 0.01D) * options.PixelsPerInch));
            RasterCanvas canvas = new(width, height, options.Supersampling, options.BackgroundColor, ResolveTextFont(options));
            canvas.Scale = options.PixelsPerInch * options.Supersampling;

            foreach (VisioShape shape in page.Shapes) {
                DrawShape(canvas, page, shape, options);
            }

            VisioRenderLabelLayout? labelLayout = options.ResolveConnectorLabelOverlaps
                ? VisioRenderLabelLayout.Create(page)
                : null;
            foreach (VisioConnector connector in page.Connectors) {
                DrawConnector(canvas, page, connector, options, labelLayout);
            }

            return EncodePngRgba(width, height, canvas.Resolve());
        }

        private static OfficeTrueTypeFont? ResolveTextFont(VisioPngSaveOptions options) {
            if (!string.IsNullOrWhiteSpace(options.FontFilePath)) {
                OfficeTrueTypeFont? configured = OfficeTrueTypeFont.TryLoadFromPath(options.FontFilePath, options.FontCollectionIndex, options.FontFaceName);
                if (configured != null) {
                    return configured;
                }
            }

            return null;
        }

        private static void DrawShape(RasterCanvas canvas, VisioPage page, VisioShape shape, VisioPngSaveOptions options) {
            string kind = VisioShapeGeometry.ResolveRenderKind(shape);
            if (VisioShapeGeometry.TryGetRenderClosedPaths(shape, out List<VisioShapeGeometryPath> preservedPaths)) {
                foreach (VisioShapeGeometryPath preservedPath in preservedPaths) {
                    List<(double X, double Y)> points = new();
                    for (int i = 0; i < preservedPath.Points.Count; i++) {
                        (double px, double py) = GetPagePoint(shape, preservedPath.Points[i].X, preservedPath.Points[i].Y);
                        points.Add(ToRaster(page, px, py, canvas.Scale));
                    }

                    Color stroke = preservedPath.NoLine || !HasVisibleLine(shape) ? Color.Transparent : shape.LineColor;
                    double strokeWidth = Math.Max(shape.LineWeight * canvas.Scale, canvas.Supersampling);
                    if (preservedPath.IsClosed) {
                        canvas.FillPolygon(points, preservedPath.NoFill || shape.FillPattern == 0 ? Color.Transparent : shape.FillColor);
                        canvas.StrokePolygon(points, stroke, strokeWidth, shape.LinePattern != 1);
                    } else {
                        canvas.StrokePolyline(points, stroke, strokeWidth, shape.LinePattern != 1);
                    }
                }
            } else if (kind == "ellipse" || kind == "circle") {
                (double centerX, double centerY) = GetPagePoint(shape, shape.LocPinX, shape.LocPinY);
                (double cx, double cy) = ToRaster(page, centerX, centerY, canvas.Scale);
                canvas.DrawEllipse(
                    cx,
                    cy,
                    Math.Abs(shape.Width * canvas.Scale / 2D),
                    Math.Abs(shape.Height * canvas.Scale / 2D),
                    shape.FillPattern == 0 ? Color.Transparent : shape.FillColor,
                    HasVisibleLine(shape) ? shape.LineColor : Color.Transparent,
                    Math.Max(shape.LineWeight * canvas.Scale, canvas.Supersampling),
                    shape.LinePattern != 1,
                    ToRasterRotation(shape.Angle),
                    cx,
                    cy);
            } else if (kind == "database") {
                DrawDatabaseShape(canvas, page, shape);
            } else {
                List<(double X, double Y)> local = VisioShapeGeometry.GetBuiltinClosedPath(shape, kind);
                List<(double X, double Y)> points = new();
                for (int i = 0; i < local.Count; i++) {
                    (double px, double py) = GetPagePoint(shape, local[i].X, local[i].Y);
                    points.Add(ToRaster(page, px, py, canvas.Scale));
                }

                canvas.FillPolygon(points, shape.FillPattern == 0 ? Color.Transparent : shape.FillColor);
                canvas.StrokePolygon(points, HasVisibleLine(shape) ? shape.LineColor : Color.Transparent, Math.Max(shape.LineWeight * canvas.Scale, canvas.Supersampling), shape.LinePattern != 1);
            }

            if (options.RenderStencilArtwork) {
                if (!DrawPackagePreviewArtwork(canvas, page, shape)) {
                    DrawStencilArtwork(canvas, page, shape);
                }
            }

            if (options.RenderText && !string.IsNullOrEmpty(shape.Text)) {
                VisioTextStyle? style = shape.TextStyle;
                double textWidth = Math.Max(0.05D, style?.TextWidth ?? shape.Width);
                double textHeight = Math.Max(0.05D, style?.TextHeight ?? shape.Height);
                double localX = (style?.TextPinX ?? shape.Width / 2D) + (textWidth / 2D) - (style?.TextLocPinX ?? textWidth / 2D);
                double localY = (style?.TextPinY ?? shape.Height / 2D) + (textHeight / 2D) - (style?.TextLocPinY ?? textHeight / 2D);
                (double textX, double textY) = GetPagePoint(shape, localX, localY);
                (double x, double y) = ToRaster(page, textX, textY, canvas.Scale);
                double horizontalMargins = (style?.LeftMargin ?? 0.05D) + (style?.RightMargin ?? 0.05D);
                double verticalMargins = (style?.TopMargin ?? 0.03D) + (style?.BottomMargin ?? 0.03D);
                DrawText(
                    canvas,
                    shape.Text!,
                    x,
                    y,
                    style,
                    10D,
                    Math.Max(canvas.Supersampling * 12D, (textWidth - horizontalMargins) * canvas.Scale),
                    Math.Max(canvas.Supersampling * 8D, (textHeight - verticalMargins) * canvas.Scale),
                    ToRasterRotation(shape.Angle + (style?.TextAngle ?? 0D)),
                    false);
            }

            foreach (VisioShape child in shape.Children) {
                DrawShape(canvas, page, child, options);
            }
        }

        private static bool HasVisibleLine(VisioShape shape) =>
            shape.LinePattern != 0 && shape.LineWeight > 0D && shape.LineColor.A > 0;

        private static void DrawDatabaseShape(RasterCanvas canvas, VisioPage page, VisioShape shape) {
            double capHeight = Math.Min(shape.Height * 0.18D, shape.Width * 0.16D);
            double midX = shape.Width / 2D;
            (double topX, double topY) = ToRasterPoint(page, shape, midX, shape.Height - capHeight, canvas.Scale);
            (double bottomX, double bottomY) = ToRasterPoint(page, shape, midX, capHeight, canvas.Scale);
            double radiusX = Math.Max(0.5D, shape.Width * canvas.Scale / 2D);
            double radiusY = Math.Max(0.5D, capHeight * canvas.Scale);
            Color fill = shape.FillPattern == 0 ? Color.Transparent : shape.FillColor;
            Color stroke = HasVisibleLine(shape) ? shape.LineColor : Color.Transparent;
            double strokeWidth = Math.Max(shape.LineWeight * canvas.Scale, canvas.Supersampling);
            bool dashed = shape.LinePattern != 1;

            List<(double X, double Y)> body = new() {
                ToRasterPoint(page, shape, 0D, capHeight, canvas.Scale),
                ToRasterPoint(page, shape, 0D, shape.Height - capHeight, canvas.Scale),
                ToRasterPoint(page, shape, shape.Width, shape.Height - capHeight, canvas.Scale),
                ToRasterPoint(page, shape, shape.Width, capHeight, canvas.Scale)
            };

            canvas.FillPolygon(body, fill);
            double rasterRotation = ToRasterRotation(shape.Angle);
            canvas.DrawEllipse(bottomX, bottomY, radiusX, radiusY, fill, Color.Transparent, strokeWidth, dashed, rasterRotation, bottomX, bottomY);
            canvas.DrawEllipse(topX, topY, radiusX, radiusY, fill, Color.Transparent, strokeWidth, dashed, rasterRotation, topX, topY);
            if (stroke.A == 0) {
                return;
            }

            canvas.StrokePolyline(
                new[] {
                    ToRasterPoint(page, shape, 0D, capHeight, canvas.Scale),
                    ToRasterPoint(page, shape, 0D, shape.Height - capHeight, canvas.Scale)
                },
                stroke,
                strokeWidth,
                dashed);
            canvas.StrokePolyline(
                new[] {
                    ToRasterPoint(page, shape, shape.Width, capHeight, canvas.Scale),
                    ToRasterPoint(page, shape, shape.Width, shape.Height - capHeight, canvas.Scale)
                },
                stroke,
                strokeWidth,
                dashed);
            canvas.DrawEllipse(bottomX, bottomY, radiusX, radiusY, Color.Transparent, stroke, strokeWidth, dashed, rasterRotation, bottomX, bottomY);
            canvas.DrawEllipse(topX, topY, radiusX, radiusY, Color.Transparent, stroke, strokeWidth, dashed, rasterRotation, topX, topY);
        }

        private static void DrawConnector(RasterCanvas canvas, VisioPage page, VisioConnector connector, VisioPngSaveOptions options, VisioRenderLabelLayout? labelLayout) {
            List<(double X, double Y)> pagePoints = GetConnectorPoints(connector);
            List<(double X, double Y)> points = new();
            for (int i = 0; i < pagePoints.Count; i++) {
                points.Add(ToRaster(page, pagePoints[i].X, pagePoints[i].Y, canvas.Scale));
            }

            bool visibleLine = connector.LinePattern != 0 && connector.LineWeight > 0D && connector.LineColor.A > 0;
            double weight = Math.Max(connector.LineWeight * canvas.Scale, canvas.Supersampling);
            canvas.StrokePolyline(points, visibleLine ? connector.LineColor : Color.Transparent, weight, connector.LinePattern != 1);

            if (visibleLine && connector.BeginArrow.HasValue && connector.BeginArrow.Value != EndArrow.None && TryGetArrowSegment(points, fromStart: true, out (double X, double Y) beginTip, out (double X, double Y) beginFrom)) {
                DrawArrow(canvas, beginTip, beginFrom, connector.LineColor, weight);
            }

            if (visibleLine && connector.EndArrow.HasValue && connector.EndArrow.Value != EndArrow.None && TryGetArrowSegment(points, fromStart: false, out (double X, double Y) endTip, out (double X, double Y) endFrom)) {
                DrawArrow(canvas, endTip, endFrom, connector.LineColor, weight);
            }

            if (options.RenderConnectorLabels && !string.IsNullOrEmpty(connector.Label)) {
                VisioRenderConnectorLabelPlacement label = labelLayout?.Resolve(connector, pagePoints) ?? ResolveConnectorLabel(connector, pagePoints);
                (double x, double y) = ToRaster(page, label.X, label.Y, canvas.Scale);
                double maxWidth = label.Width * canvas.Scale;
                double maxHeight = label.Height * canvas.Scale;
                DrawText(canvas, connector.Label!, x, y, connector.TextStyle, 9D, maxWidth, maxHeight, 0D, true);
            }
        }

        private static void DrawArrow(RasterCanvas canvas, (double X, double Y) tip, (double X, double Y) from, Color color, double weight) {
            double angle = Math.Atan2(tip.Y - from.Y, tip.X - from.X);
            double length = Math.Max(weight * 4D, canvas.Supersampling * 8D);
            double wing = Math.PI / 7D;
            List<(double X, double Y)> arrow = new() {
                tip,
                (tip.X - Math.Cos(angle - wing) * length, tip.Y - Math.Sin(angle - wing) * length),
                (tip.X - Math.Cos(angle + wing) * length, tip.Y - Math.Sin(angle + wing) * length)
            };
            canvas.FillPolygon(arrow, color);
        }

        private static bool TryGetArrowSegment(
            IReadOnlyList<(double X, double Y)> points,
            bool fromStart,
            out (double X, double Y) tip,
            out (double X, double Y) from) {
            if (points.Count < 2) {
                tip = default;
                from = default;
                return false;
            }

            if (fromStart) {
                tip = points[0];
                for (int i = 1; i < points.Count; i++) {
                    if (Distance(tip, points[i]) > 1e-6D) {
                        from = points[i];
                        return true;
                    }
                }
            } else {
                tip = points[points.Count - 1];
                for (int i = points.Count - 2; i >= 0; i--) {
                    if (Distance(tip, points[i]) > 1e-6D) {
                        from = points[i];
                        return true;
                    }
                }
            }

            from = default;
            return false;
        }

        private static void DrawText(
            RasterCanvas canvas,
            string text,
            double centerX,
            double centerY,
            VisioTextStyle? style,
            double defaultSize,
            double maxWidth,
            double maxHeight,
            double rotateRadians,
            bool drawLabelBackground) {
            double pointSize = style?.Size ?? defaultSize;
            double pixelHeight = Math.Max(canvas.Supersampling * 7D, pointSize * canvas.Scale / 72D);
            Color color = style?.Color ?? Color.FromRgb(17, 24, 39);

            string[] lines = WrapText(canvas, text, pixelHeight, maxWidth);
            double lineHeight = pixelHeight * 1.25D;
            double measuredWidth = MeasureMaxLineWidth(canvas, lines, pixelHeight);
            double measuredHeight = Math.Max(pixelHeight, ((lines.Length - 1) * lineHeight) + pixelHeight);
            double scaleDown = Math.Min(1D, Math.Min(maxWidth / Math.Max(measuredWidth, 1D), maxHeight / Math.Max(measuredHeight, 1D)));
            if (scaleDown < 0.98D) {
                pixelHeight = Math.Max(canvas.Supersampling * 5D, pixelHeight * scaleDown);
                lines = WrapText(canvas, text, pixelHeight, maxWidth);
                lineHeight = pixelHeight * 1.25D;
                measuredWidth = MeasureMaxLineWidth(canvas, lines, pixelHeight);
                measuredHeight = Math.Max(pixelHeight, ((lines.Length - 1) * lineHeight) + pixelHeight);
            }

            double top;
            switch (style?.VerticalAlignment) {
                case VisioTextVerticalAlignment.Top:
                    top = centerY - (maxHeight / 2D);
                    break;
                case VisioTextVerticalAlignment.Bottom:
                    top = centerY + (maxHeight / 2D) - measuredHeight;
                    break;
                default:
                    top = centerY - (measuredHeight / 2D);
                    break;
            }

            double anchorX = ResolveTextAnchorX(centerX, maxWidth, style?.HorizontalAlignment);
            Color? backgroundColor = ResolveTextBackground(style, drawLabelBackground);
            if (backgroundColor.HasValue) {
                double padX = Math.Max(canvas.Supersampling * 3D, pixelHeight * 0.22D);
                double padY = Math.Max(canvas.Supersampling * 2D, pixelHeight * 0.16D);
                double backgroundLeft = GetAlignedTextLeft(anchorX, measuredWidth, style?.HorizontalAlignment) - padX;
                double backgroundTop = top - padY;
                double backgroundWidth = measuredWidth + (padX * 2D);
                double backgroundHeight = measuredHeight + (padY * 2D);
                if (Math.Abs(rotateRadians) < TextRotationEpsilon) {
                    canvas.FillRectangle(backgroundLeft, backgroundTop, backgroundWidth, backgroundHeight, backgroundColor.Value);
                } else {
                    canvas.FillPolygon(new[] {
                        RotateTextPoint((backgroundLeft, backgroundTop), centerX, centerY, rotateRadians),
                        RotateTextPoint((backgroundLeft + backgroundWidth, backgroundTop), centerX, centerY, rotateRadians),
                        RotateTextPoint((backgroundLeft + backgroundWidth, backgroundTop + backgroundHeight), centerX, centerY, rotateRadians),
                        RotateTextPoint((backgroundLeft, backgroundTop + backgroundHeight), centerX, centerY, rotateRadians)
                    }, backgroundColor.Value);
                }
            }

            for (int i = 0; i < lines.Length; i++) {
                double lineTop = top + (i * lineHeight);
                canvas.DrawTextLine(lines[i], anchorX, lineTop, pixelHeight, color, style?.Bold == true, style?.Italic == true, style?.HorizontalAlignment, rotateRadians, centerX, centerY);
                if (style?.Underline == true) {
                    double lineWidth = canvas.MeasureText(lines[i], pixelHeight);
                    double underlineY = lineTop + (pixelHeight * 0.92D);
                    double underlineLeft = GetAlignedTextLeft(anchorX, lineWidth, style.HorizontalAlignment);
                    double underlineWeight = Math.Max(canvas.Supersampling, pixelHeight / 16D);
                    (double X, double Y) underlineStart = RotateTextPoint((underlineLeft, underlineY), centerX, centerY, rotateRadians);
                    (double X, double Y) underlineEnd = RotateTextPoint((underlineLeft + lineWidth, underlineY), centerX, centerY, rotateRadians);
                    canvas.StrokePolyline(
                        new[] { underlineStart, underlineEnd },
                        color,
                        underlineWeight,
                        dashed: false);
                }
            }
        }

        private static string[] WrapText(RasterCanvas canvas, string text, double pixelHeight, double maxWidth) {
            string[] sourceLines = text.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n');
            List<string> output = new();
            foreach (string sourceLine in sourceLines) {
                string line = sourceLine.Trim();
                if (line.Length == 0) {
                    output.Add(string.Empty);
                    continue;
                }

                string[] words = line.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                string current = string.Empty;
                for (int i = 0; i < words.Length; i++) {
                    string word = words[i];
                    if (canvas.MeasureText(word, pixelHeight) > maxWidth) {
                        if (current.Length > 0) {
                            output.Add(current);
                            current = string.Empty;
                        }

                        foreach (string part in BreakWord(canvas, word, pixelHeight, maxWidth)) {
                            output.Add(part);
                        }

                        continue;
                    }

                    string candidate = current.Length == 0 ? word : current + " " + word;
                    if (current.Length > 0 && canvas.MeasureText(candidate, pixelHeight) > maxWidth) {
                        output.Add(current);
                        current = word;
                    } else {
                        current = candidate;
                    }
                }

                if (current.Length > 0) {
                    output.Add(current);
                }
            }

            return output.Count == 0 ? new[] { string.Empty } : output.ToArray();
        }

        private static IEnumerable<string> BreakWord(RasterCanvas canvas, string word, double pixelHeight, double maxWidth) {
            StringBuilder part = new();
            foreach (char c in word) {
                string candidate = part.ToString() + c;
                if (part.Length > 0 && canvas.MeasureText(candidate, pixelHeight) > maxWidth) {
                    yield return part.ToString();
                    part.Clear();
                }

                part.Append(c);
            }

            if (part.Length > 0) {
                yield return part.ToString();
            }
        }

        private const double TextRotationEpsilon = 1e-9;

        private static (double X, double Y) RotateTextPoint((double X, double Y) point, double centerX, double centerY, double radians) {
            if (Math.Abs(radians) < TextRotationEpsilon) return point;
            double cos = Math.Cos(-radians);
            double sin = Math.Sin(-radians);
            double dx = point.X - centerX;
            double dy = point.Y - centerY;
            return (centerX + (dx * cos) - (dy * sin), centerY + (dx * sin) + (dy * cos));
        }

        private static double MeasureMaxLineWidth(RasterCanvas canvas, IReadOnlyList<string> lines, double pixelHeight) {
            double max = 0D;
            for (int i = 0; i < lines.Count; i++) {
                max = Math.Max(max, canvas.MeasureText(lines[i], pixelHeight));
            }

            return max;
        }

        private static double ResolveTextAnchorX(double centerX, double maxWidth, VisioTextHorizontalAlignment? alignment) {
            switch (alignment) {
                case VisioTextHorizontalAlignment.Left:
                    return centerX - (maxWidth / 2D);
                case VisioTextHorizontalAlignment.Right:
                    return centerX + (maxWidth / 2D);
                default:
                    return centerX;
            }
        }

        private static double GetAlignedTextLeft(double anchorX, double width, VisioTextHorizontalAlignment? alignment) {
            switch (alignment) {
                case VisioTextHorizontalAlignment.Left:
                    return anchorX;
                case VisioTextHorizontalAlignment.Right:
                    return anchorX - width;
                default:
                    return anchorX - (width / 2D);
            }
        }

        private static Color? ResolveTextBackground(VisioTextStyle? style, bool drawLabelBackground) {
            if (style?.BackgroundColor.HasValue == true) {
                return ApplyBackgroundTransparency(style.BackgroundColor.Value, style.BackgroundTransparency);
            }

            return drawLabelBackground ? Color.FromRgba(255, 255, 255, 230) : null;
        }

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
                RotateTextPoint((x + (size * offsetX), y + (size * offsetY)), x, y, rasterRotation);
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

        private static bool DrawPackagePreviewArtwork(RasterCanvas canvas, VisioPage page, VisioShape shape) {
            if (!VisioPackagePreviewArtwork.TryGetPng(shape, out VisioPreviewImage image) ||
                !PngRaster.TryDecode(image.Data, out PngRaster? raster) ||
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
            double imageAspect = (double)raster.Width / raster.Height;
            double targetAspect = targetWidth / targetHeight;
            double drawWidth = targetWidth;
            double drawHeight = targetHeight;
            if (imageAspect > targetAspect) {
                drawHeight = targetWidth / imageAspect;
            } else if (imageAspect < targetAspect) {
                drawWidth = targetHeight * imageAspect;
            }

            canvas.DrawImage(
                raster,
                centerX - (drawWidth / 2D),
                centerY - (drawHeight / 2D),
                drawWidth,
                drawHeight,
                ToRasterRotation(shape.Angle),
                centerX,
                centerY);
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

        private static void StrokeLine(RasterCanvas canvas, double x1, double y1, double x2, double y2, Color color, double width) =>
            canvas.StrokePolyline(new[] { (x1, y1), (x2, y2) }, color, width, dashed: false);

        private static void StrokeRect(RasterCanvas canvas, double x, double y, double width, double height, Color color, double stroke) =>
            StrokePolyline(canvas, new[] { (x, y), (x + width, y), (x + width, y + height), (x, y + height), (x, y) }, color, stroke);

        private static void StrokeEllipse(RasterCanvas canvas, double x, double y, double rx, double ry, Color color, double stroke, double rotationRadians = 0D, double rotationCenterX = 0D, double rotationCenterY = 0D) {
            if (Math.Abs(rotationRadians) <= 1e-9) {
                canvas.DrawEllipse(x, y, rx, ry, Color.Transparent, color, stroke);
                return;
            }

            List<(double X, double Y)> points = new();
            for (int i = 0; i <= 36; i++) {
                double angle = (Math.PI * 2D) * i / 36D;
                (double X, double Y) point = (x + (Math.Cos(angle) * rx), y + (Math.Sin(angle) * ry));
                points.Add(RotateTextPoint(point, rotationCenterX, rotationCenterY, rotationRadians));
            }

            StrokePolyline(canvas, points, color, stroke);
        }

        private static void StrokeArc(RasterCanvas canvas, double x, double y, double rx, double ry, double startDegrees, double endDegrees, Color color, double stroke, double rotationRadians = 0D, double rotationCenterX = 0D, double rotationCenterY = 0D) {
            List<(double X, double Y)> points = new();
            for (int i = 0; i <= 18; i++) {
                double angle = (startDegrees + ((endDegrees - startDegrees) * i / 18D)) * Math.PI / 180D;
                (double X, double Y) point = (x + Math.Cos(angle) * rx, y + Math.Sin(angle) * ry);
                if (Math.Abs(rotationRadians) > 1e-9) {
                    point = RotateTextPoint(point, rotationCenterX, rotationCenterY, rotationRadians);
                }

                points.Add(point);
            }

            StrokePolyline(canvas, points, color, stroke);
        }

        private static void StrokePolyline(RasterCanvas canvas, IReadOnlyList<(double X, double Y)> points, Color color, double stroke) =>
            canvas.StrokePolyline(points, color, stroke, dashed: false);

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
            List<(double X, double Y)> points = new() { (startX, startY) };
            if (connector.Waypoints.Count > 0) {
                foreach (VisioConnectorWaypoint waypoint in connector.Waypoints) {
                    points.Add((waypoint.X, waypoint.Y));
                }
            } else if (connector.Kind == ConnectorKind.RightAngle) {
                points.Add((startX, endY));
            }

            points.Add((endX, endY));
            return points;
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
            (double x, double y) = InterpolatePath(points, position);
            return (x + (placement?.OffsetX ?? 0D), y + (placement?.OffsetY ?? 0D));
        }

        private static VisioRenderConnectorLabelPlacement ResolveConnectorLabel(VisioConnector connector, IReadOnlyList<(double X, double Y)> points) {
            (double x, double y) = ResolveConnectorLabelPoint(connector, points);
            VisioConnectorLabelPlacement? placement = connector.LabelPlacement;
            double width = Math.Max(0.6D, connector.TextStyle?.TextWidth ?? placement?.Width ?? 1.35D);
            double height = Math.Max(0.18D, connector.TextStyle?.TextHeight ?? placement?.Height ?? 0.34D);
            return new VisioRenderConnectorLabelPlacement(x, y, width, height, adjusted: false);
        }

        private static (double X, double Y) InterpolatePath(IReadOnlyList<(double X, double Y)> points, double position) {
            if (points.Count == 0) return (0D, 0D);
            if (points.Count == 1) return points[0];

            double total = 0D;
            for (int i = 1; i < points.Count; i++) {
                total += Distance(points[i - 1], points[i]);
            }

            if (total <= 0D) return points[0];
            double target = total * position;
            double traversed = 0D;
            for (int i = 1; i < points.Count; i++) {
                double segment = Distance(points[i - 1], points[i]);
                if (traversed + segment >= target) {
                    double t = segment <= 0D ? 0D : (target - traversed) / segment;
                    return (
                        points[i - 1].X + ((points[i].X - points[i - 1].X) * t),
                        points[i - 1].Y + ((points[i].Y - points[i - 1].Y) * t));
                }

                traversed += segment;
            }

            return points[points.Count - 1];
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

        private static (double X, double Y) ToRaster(VisioPage page, double x, double y, double scale) =>
            (x * scale, (page.Height - y) * scale);

        private static (double X, double Y) ToRasterPoint(VisioPage page, VisioShape shape, double x, double y, double scale) {
            (double pageX, double pageY) = GetPagePoint(shape, x, y);
            return ToRaster(page, pageX, pageY, scale);
        }

        private static double Distance((double X, double Y) a, (double X, double Y) b) {
            double dx = b.X - a.X;
            double dy = b.Y - a.Y;
            return Math.Sqrt((dx * dx) + (dy * dy));
        }

        private static byte[] EncodePngRgba(int width, int height, byte[] rgba) {
            byte[] scanlines = new byte[height * (1 + width * 4)];
            int source = 0;
            int target = 0;
            for (int y = 0; y < height; y++) {
                scanlines[target++] = 0;
                Buffer.BlockCopy(rgba, source, scanlines, target, width * 4);
                source += width * 4;
                target += width * 4;
            }

            using MemoryStream ms = new();
            ms.Write(PngSignature, 0, PngSignature.Length);
            byte[] ihdr = new byte[13];
            WriteBigEndianInt32(ihdr, 0, width);
            WriteBigEndianInt32(ihdr, 4, height);
            ihdr[8] = 8;
            ihdr[9] = 6;
            WriteChunk(ms, "IHDR", ihdr);
            WriteChunk(ms, "IDAT", DeflateZlib(scanlines));
            WriteChunk(ms, "IEND", Array.Empty<byte>());
            return ms.ToArray();
        }

        private static byte[] DeflateZlib(byte[] data) {
            using MemoryStream ms = new();
            ms.WriteByte(0x78);
            ms.WriteByte(0x9C);
            using (DeflateStream deflate = new(ms, CompressionLevel.Optimal, leaveOpen: true)) {
                deflate.Write(data, 0, data.Length);
            }

            uint adler = Adler32(data);
            ms.WriteByte((byte)((adler >> 24) & 0xFF));
            ms.WriteByte((byte)((adler >> 16) & 0xFF));
            ms.WriteByte((byte)((adler >> 8) & 0xFF));
            ms.WriteByte((byte)(adler & 0xFF));
            return ms.ToArray();
        }

        private static uint Adler32(byte[] data) {
            const uint mod = 65521;
            uint a = 1;
            uint b = 0;
            for (int i = 0; i < data.Length; i++) {
                a = (a + data[i]) % mod;
                b = (b + a) % mod;
            }

            return (b << 16) | a;
        }

        private static void WriteChunk(Stream stream, string type, byte[] data) {
            byte[] typeBytes = Encoding.ASCII.GetBytes(type);
            byte[] length = new byte[4];
            WriteBigEndianInt32(length, 0, data.Length);
            stream.Write(length, 0, length.Length);
            stream.Write(typeBytes, 0, typeBytes.Length);
            stream.Write(data, 0, data.Length);

            uint crc = Crc32(typeBytes, data);
            byte[] crcBytes = new byte[4];
            WriteBigEndianInt32(crcBytes, 0, unchecked((int)crc));
            stream.Write(crcBytes, 0, crcBytes.Length);
        }

        private static uint Crc32(byte[] type, byte[] data) {
            uint crc = 0xFFFFFFFF;
            for (int i = 0; i < type.Length; i++) crc = UpdateCrc(crc, type[i]);
            for (int i = 0; i < data.Length; i++) crc = UpdateCrc(crc, data[i]);
            return crc ^ 0xFFFFFFFF;
        }

        private static uint UpdateCrc(uint crc, byte value) {
            crc ^= value;
            for (int i = 0; i < 8; i++) {
                crc = (crc & 1) != 0 ? 0xEDB88320 ^ (crc >> 1) : crc >> 1;
            }

            return crc;
        }

        private static void WriteBigEndianInt32(byte[] bytes, int offset, int value) {
            bytes[offset] = (byte)((value >> 24) & 0xFF);
            bytes[offset + 1] = (byte)((value >> 16) & 0xFF);
            bytes[offset + 2] = (byte)((value >> 8) & 0xFF);
            bytes[offset + 3] = (byte)(value & 0xFF);
        }

        private sealed class PngRaster {
            private static readonly byte[] Signature = { 137, 80, 78, 71, 13, 10, 26, 10 };
            private readonly byte[] _pixels;

            private PngRaster(int width, int height, byte[] pixels) {
                Width = width;
                Height = height;
                _pixels = pixels;
            }

            internal int Width { get; }

            internal int Height { get; }

            internal Color GetPixel(int x, int y) {
                int offset = ((y * Width) + x) * 4;
                return Color.FromRgba(_pixels[offset], _pixels[offset + 1], _pixels[offset + 2], _pixels[offset + 3]);
            }

            private static int ReadBigEndianInt32(byte[] bytes, int offset) =>
                (bytes[offset] << 24) | (bytes[offset + 1] << 16) | (bytes[offset + 2] << 8) | bytes[offset + 3];

            private static bool HasPngSignature(byte[] bytes) {
                if (bytes.Length < Signature.Length) {
                    return false;
                }

                for (int i = 0; i < Signature.Length; i++) {
                    if (bytes[i] != Signature[i]) {
                        return false;
                    }
                }

                return true;
            }

            internal static bool TryDecode(byte[] bytes, out PngRaster? image) {
                image = null;
                try {
                    if (!HasPngSignature(bytes)) {
                        return false;
                    }

                    int width = 0;
                    int height = 0;
                    int bitDepth = 0;
                    int colorType = 0;
                    int compressionMethod = 0;
                    int filterMethod = 0;
                    int interlaceMethod = 0;
                    byte[]? palette = null;
                    byte[]? transparency = null;
                    using MemoryStream idat = new();
                    int offset = Signature.Length;
                    while (offset + 12 <= bytes.Length) {
                        int length = ReadBigEndianInt32(bytes, offset);
                        if (length < 0 || offset + 12 + length > bytes.Length) {
                            return false;
                        }

                        string type = Encoding.ASCII.GetString(bytes, offset + 4, 4);
                        int dataOffset = offset + 8;
                        if (type == "IHDR") {
                            width = ReadBigEndianInt32(bytes, dataOffset);
                            height = ReadBigEndianInt32(bytes, dataOffset + 4);
                            bitDepth = bytes[dataOffset + 8];
                            colorType = bytes[dataOffset + 9];
                            compressionMethod = bytes[dataOffset + 10];
                            filterMethod = bytes[dataOffset + 11];
                            interlaceMethod = bytes[dataOffset + 12];
                        } else if (type == "PLTE") {
                            palette = new byte[length];
                            Buffer.BlockCopy(bytes, dataOffset, palette, 0, length);
                        } else if (type == "tRNS") {
                            transparency = new byte[length];
                            Buffer.BlockCopy(bytes, dataOffset, transparency, 0, length);
                        } else if (type == "IDAT") {
                            idat.Write(bytes, dataOffset, length);
                        } else if (type == "IEND") {
                            break;
                        }

                        offset = dataOffset + length + 4;
                    }

                    if (width <= 0 || height <= 0 || compressionMethod != 0 || filterMethod != 0 || interlaceMethod != 0 ||
                        !IsSupportedColorLayout(colorType, bitDepth, palette)) {
                        return false;
                    }

                    int bitsPerPixel = GetBitsPerPixel(colorType, bitDepth);
                    int bytesPerPixel = Math.Max(1, (bitsPerPixel + 7) / 8);
                    byte[] compressed = idat.ToArray();
                    if (compressed.Length < 6) {
                        return false;
                    }

                    using MemoryStream source = new(compressed, 2, compressed.Length - 6);
                    using DeflateStream deflate = new(source, CompressionMode.Decompress);
                    using MemoryStream inflated = new();
                    deflate.CopyTo(inflated);
                    byte[] scanlines = inflated.ToArray();
                    int stride = ((width * bitsPerPixel) + 7) / 8;
                    byte[] previous = new byte[stride];
                    byte[] current = new byte[stride];
                    byte[] rgba = new byte[width * height * 4];
                    int sourceOffset = 0;
                    for (int y = 0; y < height; y++) {
                        if (sourceOffset >= scanlines.Length) return false;
                        int filter = scanlines[sourceOffset++];
                        if (sourceOffset + stride > scanlines.Length) return false;
                        Buffer.BlockCopy(scanlines, sourceOffset, current, 0, stride);
                        sourceOffset += stride;
                        Unfilter(current, previous, bytesPerPixel, filter);
                        ExpandScanline(current, width, y, colorType, bitDepth, palette, transparency, rgba);

                        byte[] temp = previous;
                        previous = current;
                        current = temp;
                        Array.Clear(current, 0, current.Length);
                    }

                    image = new PngRaster(width, height, rgba);
                    return true;
                } catch {
                    image = null;
                    return false;
                }
            }

            private static bool IsSupportedColorLayout(int colorType, int bitDepth, byte[]? palette) {
                switch (colorType) {
                    case 0:
                        return bitDepth == 1 || bitDepth == 2 || bitDepth == 4 || bitDepth == 8 || bitDepth == 16;
                    case 2:
                    case 4:
                    case 6:
                        return bitDepth == 8 || bitDepth == 16;
                    case 3:
                        return (bitDepth == 1 || bitDepth == 2 || bitDepth == 4 || bitDepth == 8) &&
                               palette != null &&
                               palette.Length >= 3 &&
                               palette.Length % 3 == 0;
                    default:
                        return false;
                }
            }

            private static int GetBitsPerPixel(int colorType, int bitDepth) {
                switch (colorType) {
                    case 0:
                    case 3:
                        return bitDepth;
                    case 2:
                        return bitDepth * 3;
                    case 4:
                        return bitDepth * 2;
                    case 6:
                        return bitDepth * 4;
                    default:
                        throw new InvalidDataException("Unsupported PNG color type.");
                }
            }

            private static void ExpandScanline(
                byte[] current,
                int width,
                int y,
                int colorType,
                int bitDepth,
                byte[]? palette,
                byte[]? transparency,
                byte[] rgba) {
                for (int x = 0; x < width; x++) {
                    int targetPixel = ((y * width) + x) * 4;
                    switch (colorType) {
                        case 0:
                            ExpandGrayscale(
                                GetGrayscaleSample(current, x, bitDepth),
                                bitDepth,
                                transparency,
                                rgba,
                                targetPixel);
                            break;
                        case 2:
                            ExpandTrueColor(current, x * (bitDepth == 16 ? 6 : 3), bitDepth, transparency, rgba, targetPixel);
                            break;
                        case 3:
                            ExpandPalette(GetPackedSample(current, x, bitDepth), palette!, transparency, rgba, targetPixel);
                            break;
                        case 4:
                            ExpandGrayscaleAlpha(current, x * (bitDepth == 16 ? 4 : 2), bitDepth, rgba, targetPixel);
                            break;
                        case 6:
                            ExpandTrueColorAlpha(current, x * (bitDepth == 16 ? 8 : 4), bitDepth, rgba, targetPixel);
                            break;
                        default:
                            throw new InvalidDataException("Unsupported PNG color type.");
                    }
                }
            }

            private static void ExpandGrayscale(int sample, int bitDepth, byte[]? transparency, byte[] rgba, int targetPixel) {
                byte gray = ScaleSample(sample, bitDepth);
                rgba[targetPixel] = gray;
                rgba[targetPixel + 1] = gray;
                rgba[targetPixel + 2] = gray;
                rgba[targetPixel + 3] = IsTransparentGray(sample, transparency) ? (byte)0 : (byte)255;
            }

            private static void ExpandGrayscaleAlpha(byte[] current, int sourcePixel, int bitDepth, byte[] rgba, int targetPixel) {
                int graySample = bitDepth == 16 ? ReadBigEndianUInt16(current, sourcePixel) : current[sourcePixel];
                int alphaSample = bitDepth == 16 ? ReadBigEndianUInt16(current, sourcePixel + 2) : current[sourcePixel + 1];
                byte gray = ScaleSample(graySample, bitDepth);
                rgba[targetPixel] = gray;
                rgba[targetPixel + 1] = gray;
                rgba[targetPixel + 2] = gray;
                rgba[targetPixel + 3] = ScaleSample(alphaSample, bitDepth);
            }

            private static void ExpandTrueColor(byte[] current, int sourcePixel, int bitDepth, byte[]? transparency, byte[] rgba, int targetPixel) {
                int redSample;
                int greenSample;
                int blueSample;
                if (bitDepth == 16) {
                    redSample = ReadBigEndianUInt16(current, sourcePixel);
                    greenSample = ReadBigEndianUInt16(current, sourcePixel + 2);
                    blueSample = ReadBigEndianUInt16(current, sourcePixel + 4);
                } else {
                    redSample = current[sourcePixel];
                    greenSample = current[sourcePixel + 1];
                    blueSample = current[sourcePixel + 2];
                }

                rgba[targetPixel] = ScaleSample(redSample, bitDepth);
                rgba[targetPixel + 1] = ScaleSample(greenSample, bitDepth);
                rgba[targetPixel + 2] = ScaleSample(blueSample, bitDepth);
                rgba[targetPixel + 3] = IsTransparentRgb(redSample, greenSample, blueSample, transparency) ? (byte)0 : (byte)255;
            }

            private static void ExpandTrueColorAlpha(byte[] current, int sourcePixel, int bitDepth, byte[] rgba, int targetPixel) {
                if (bitDepth == 16) {
                    rgba[targetPixel] = ScaleSample(ReadBigEndianUInt16(current, sourcePixel), bitDepth);
                    rgba[targetPixel + 1] = ScaleSample(ReadBigEndianUInt16(current, sourcePixel + 2), bitDepth);
                    rgba[targetPixel + 2] = ScaleSample(ReadBigEndianUInt16(current, sourcePixel + 4), bitDepth);
                    rgba[targetPixel + 3] = ScaleSample(ReadBigEndianUInt16(current, sourcePixel + 6), bitDepth);
                } else {
                    Buffer.BlockCopy(current, sourcePixel, rgba, targetPixel, 4);
                }
            }

            private static void ExpandPalette(int index, byte[] palette, byte[]? transparency, byte[] rgba, int targetPixel) {
                int paletteOffset = index * 3;
                if (paletteOffset + 2 >= palette.Length) {
                    throw new InvalidDataException("PNG palette index is outside PLTE.");
                }

                rgba[targetPixel] = palette[paletteOffset];
                rgba[targetPixel + 1] = palette[paletteOffset + 1];
                rgba[targetPixel + 2] = palette[paletteOffset + 2];
                rgba[targetPixel + 3] = transparency != null && index < transparency.Length ? transparency[index] : (byte)255;
            }

            private static int GetPackedSample(byte[] current, int x, int bitDepth) {
                if (bitDepth == 8) {
                    return current[x];
                }

                int samplesPerByte = 8 / bitDepth;
                int shift = (samplesPerByte - 1 - (x % samplesPerByte)) * bitDepth;
                int mask = (1 << bitDepth) - 1;
                return (current[x / samplesPerByte] >> shift) & mask;
            }

            private static int GetGrayscaleSample(byte[] current, int x, int bitDepth) {
                if (bitDepth == 16) {
                    return ReadBigEndianUInt16(current, x * 2);
                }

                return bitDepth == 8 ? current[x] : GetPackedSample(current, x, bitDepth);
            }

            private static int ReadBigEndianUInt16(byte[] bytes, int offset) =>
                (bytes[offset] << 8) | bytes[offset + 1];

            private static byte ScaleSample(int sample, int bitDepth) {
                if (bitDepth == 8) {
                    return (byte)sample;
                }

                int max = (1 << bitDepth) - 1;
                return (byte)Math.Round(sample * 255D / max);
            }

            private static bool IsTransparentGray(int sample, byte[]? transparency) =>
                transparency != null && transparency.Length >= 2 && sample == ((transparency[0] << 8) | transparency[1]);

            private static bool IsTransparentRgb(int red, int green, int blue, byte[]? transparency) =>
                transparency != null &&
                transparency.Length >= 6 &&
                red == ((transparency[0] << 8) | transparency[1]) &&
                green == ((transparency[2] << 8) | transparency[3]) &&
                blue == ((transparency[4] << 8) | transparency[5]);

            private static void Unfilter(byte[] current, byte[] previous, int bytesPerPixel, int filter) {
                for (int i = 0; i < current.Length; i++) {
                    int left = i >= bytesPerPixel ? current[i - bytesPerPixel] : 0;
                    int up = previous[i];
                    int upLeft = i >= bytesPerPixel ? previous[i - bytesPerPixel] : 0;
                    int value = current[i];
                    switch (filter) {
                        case 0:
                            break;
                        case 1:
                            value += left;
                            break;
                        case 2:
                            value += up;
                            break;
                        case 3:
                            value += (left + up) / 2;
                            break;
                        case 4:
                            value += Paeth(left, up, upLeft);
                            break;
                        default:
                            throw new InvalidDataException("Unsupported PNG filter.");
                    }

                    current[i] = (byte)(value & 0xFF);
                }
            }

            private static int Paeth(int left, int up, int upLeft) {
                int p = left + up - upLeft;
                int pa = Math.Abs(p - left);
                int pb = Math.Abs(p - up);
                int pc = Math.Abs(p - upLeft);
                if (pa <= pb && pa <= pc) return left;
                return pb <= pc ? up : upLeft;
            }
        }

        private sealed class RasterCanvas {
            private static readonly OfficeTrueTypeFont? DefaultOutlineFont = OfficeTrueTypeFont.TryLoadDefault();
            private readonly int _width;
            private readonly int _height;
            private readonly int _renderWidth;
            private readonly int _renderHeight;
            private readonly byte[] _pixels;
            private readonly OfficeTrueTypeFont? _outlineFont;

            internal RasterCanvas(int width, int height, int supersampling, Color? background, OfficeTrueTypeFont? outlineFont) {
                _width = width;
                _height = height;
                Supersampling = supersampling;
                _renderWidth = width * supersampling;
                _renderHeight = height * supersampling;
                Scale = supersampling;
                _pixels = new byte[_renderWidth * _renderHeight * 4];
                _outlineFont = outlineFont ?? DefaultOutlineFont;

                if (background.HasValue) {
                    for (int y = 0; y < _renderHeight; y++) {
                        for (int x = 0; x < _renderWidth; x++) {
                            SetPixel(x, y, background.Value);
                        }
                    }
                }
            }

            internal double Scale { get; set; }

            internal int Supersampling { get; }

            internal void FillPolygon(IReadOnlyList<(double X, double Y)> points, Color color) {
                if (color.A == 0 || points.Count < 3) return;
                (int minX, int minY, int maxX, int maxY) = Bounds(points, 1D);
                for (int y = minY; y <= maxY; y++) {
                    for (int x = minX; x <= maxX; x++) {
                        if (ContainsPoint(points, x + 0.5D, y + 0.5D)) {
                            BlendPixel(x, y, color);
                        }
                    }
                }
            }

            internal void StrokePolygon(IReadOnlyList<(double X, double Y)> points, Color color, double width, bool dashed) {
                if (points.Count == 0) return;
                List<(double X, double Y)> closed = new(points) { points[0] };
                StrokePolyline(closed, color, width, dashed);
            }

            internal void StrokePolyline(IReadOnlyList<(double X, double Y)> points, Color color, double width, bool dashed) {
                if (color.A == 0 || points.Count < 2 || width <= 0D) return;
                for (int i = 1; i < points.Count; i++) {
                    if (dashed) {
                        StrokeDashedSegment(points[i - 1], points[i], color, width);
                    } else {
                        StrokeSegment(points[i - 1], points[i], color, width);
                    }
                }
            }

            internal void DrawEllipse(double cx, double cy, double rx, double ry, Color fill, Color stroke, double width, bool dashed = false, double rotationRadians = 0D, double rotationCenterX = 0D, double rotationCenterY = 0D) {
                if (dashed && stroke.A > 0 && width > 0D) {
                    int segments = 72;
                    List<(double X, double Y)> points = new(segments + 1);
                    for (int i = 0; i <= segments; i++) {
                        double angle = Math.PI * 2D * i / segments;
                        (double X, double Y) point = (cx + (Math.Cos(angle) * rx), cy + (Math.Sin(angle) * ry));
                        if (Math.Abs(rotationRadians) > 1e-9) {
                            point = RotateTextPoint(point, rotationCenterX, rotationCenterY, rotationRadians);
                        }

                        points.Add(point);
                    }

                    DrawEllipse(cx, cy, rx, ry, fill, Color.Transparent, 0D, dashed: false, rotationRadians, rotationCenterX, rotationCenterY);
                    StrokePolyline(points, stroke, width, dashed: true);
                    return;
                }

                double strokeHalf = width / 2D;
                double boundsRadius = Math.Max(rx, ry) + strokeHalf + 1D;
                int minX = ClampToInt(Math.Floor(cx - boundsRadius), 0, _renderWidth - 1);
                int maxX = ClampToInt(Math.Ceiling(cx + boundsRadius), 0, _renderWidth - 1);
                int minY = ClampToInt(Math.Floor(cy - boundsRadius), 0, _renderHeight - 1);
                int maxY = ClampToInt(Math.Ceiling(cy + boundsRadius), 0, _renderHeight - 1);
                double outerRx = Math.Max(rx + strokeHalf, 0.1D);
                double outerRy = Math.Max(ry + strokeHalf, 0.1D);
                double innerRx = Math.Max(rx - strokeHalf, 0.1D);
                double innerRy = Math.Max(ry - strokeHalf, 0.1D);
                for (int y = minY; y <= maxY; y++) {
                    for (int x = minX; x <= maxX; x++) {
                        (double X, double Y) local = Math.Abs(rotationRadians) > 1e-9
                            ? RotateTextPoint((x + 0.5D, y + 0.5D), rotationCenterX, rotationCenterY, -rotationRadians)
                            : (x + 0.5D, y + 0.5D);
                        double dx = local.X - cx;
                        double dy = local.Y - cy;
                        double fillMetric = (dx * dx / (rx * rx)) + (dy * dy / (ry * ry));
                        if (fill.A > 0 && fillMetric <= 1D) {
                            BlendPixel(x, y, fill);
                            continue;
                        }

                        double outer = (dx * dx / (outerRx * outerRx)) + (dy * dy / (outerRy * outerRy));
                        double inner = (dx * dx / (innerRx * innerRx)) + (dy * dy / (innerRy * innerRy));
                        if (stroke.A > 0 && outer <= 1D && inner >= 1D) {
                            BlendPixel(x, y, stroke);
                        }
                    }
                }
            }

            internal void DrawImage(PngRaster image, double x, double y, double width, double height) =>
                DrawImage(image, x, y, width, height, 0D, x + (width / 2D), y + (height / 2D));

            internal void DrawImage(PngRaster image, double x, double y, double width, double height, double rotationRadians, double rotationCenterX, double rotationCenterY) {
                if (image.Width <= 0 || image.Height <= 0 || width <= 0D || height <= 0D) return;
                (double X, double Y) topLeft = RotateTextPoint((x, y), rotationCenterX, rotationCenterY, rotationRadians);
                (double X, double Y) topRight = RotateTextPoint((x + width, y), rotationCenterX, rotationCenterY, rotationRadians);
                (double X, double Y) bottomRight = RotateTextPoint((x + width, y + height), rotationCenterX, rotationCenterY, rotationRadians);
                (double X, double Y) bottomLeft = RotateTextPoint((x, y + height), rotationCenterX, rotationCenterY, rotationRadians);
                double minPointX = Math.Min(Math.Min(topLeft.X, topRight.X), Math.Min(bottomRight.X, bottomLeft.X));
                double maxPointX = Math.Max(Math.Max(topLeft.X, topRight.X), Math.Max(bottomRight.X, bottomLeft.X));
                double minPointY = Math.Min(Math.Min(topLeft.Y, topRight.Y), Math.Min(bottomRight.Y, bottomLeft.Y));
                double maxPointY = Math.Max(Math.Max(topLeft.Y, topRight.Y), Math.Max(bottomRight.Y, bottomLeft.Y));
                int minX = ClampToInt(Math.Floor(minPointX), 0, _renderWidth - 1);
                int maxX = ClampToInt(Math.Ceiling(maxPointX), 0, _renderWidth - 1);
                int minY = ClampToInt(Math.Floor(minPointY), 0, _renderHeight - 1);
                int maxY = ClampToInt(Math.Ceiling(maxPointY), 0, _renderHeight - 1);
                for (int py = minY; py <= maxY; py++) {
                    for (int px = minX; px <= maxX; px++) {
                        (double X, double Y) local = RotateTextPoint((px + 0.5D, py + 0.5D), rotationCenterX, rotationCenterY, -rotationRadians);
                        double u = (local.X - x) / width;
                        double v = (local.Y - y) / height;
                        if (u < 0D || u >= 1D || v < 0D || v >= 1D) {
                            continue;
                        }

                        int sourceX = ClampToInt(Math.Floor(u * image.Width), 0, image.Width - 1);
                        int sourceY = ClampToInt(Math.Floor(v * image.Height), 0, image.Height - 1);
                        Color color = image.GetPixel(sourceX, sourceY);
                        if (color.A > 0) {
                            BlendPixel(px, py, color);
                        }
                    }
                }
            }

            internal double MeasureText(string text, double height) {
                if (string.IsNullOrEmpty(text)) return 0D;
                if (_outlineFont != null) {
                    return _outlineFont.Measure(text, Math.Max(1D, height));
                }

                return MeasureStrokeText(text, height);
            }

            private double MeasureStrokeText(string text, double height) {
                if (string.IsNullOrEmpty(text)) return 0D;
                double cell = Math.Max(1D, height / 7D);
                double gap = cell * 0.9D;
                double width = 0D;
                foreach (char c in text) {
                    width += GlyphWidth(c) * cell + gap;
                }

                return width > 0D ? width - gap : 0D;
            }

            internal void DrawTextLine(string text, double anchorX, double top, double height, Color color, bool bold, bool italic, VisioTextHorizontalAlignment? alignment, double rotationRadians, double rotationCenterX, double rotationCenterY) {
                if (string.IsNullOrEmpty(text) || color.A == 0) return;
                if (_outlineFont != null) {
                    double width = MeasureText(text, height);
                    double x = anchorX;
                    if (alignment == VisioTextHorizontalAlignment.Right) {
                        x -= width;
                    } else if (alignment != VisioTextHorizontalAlignment.Left) {
                        x -= width / 2D;
                    }

                    double fontHeight = Math.Max(1D, height);
                    double bottom = top + fontHeight;
                    FillContours(TransformContours(_outlineFont.GetTextContours(text, x, top, fontHeight), bottom, italic, rotationRadians, rotationCenterX, rotationCenterY), color);
                    if (bold) {
                        FillContours(TransformContours(_outlineFont.GetTextContours(text, x + Math.Max(1D, fontHeight / 22D), top, fontHeight), bottom, italic, rotationRadians, rotationCenterX, rotationCenterY), color);
                    }

                    return;
                }

                DrawStrokeText(text, anchorX, top + (height / 2D), height, color, bold, italic, alignment, rotationRadians, rotationCenterX, rotationCenterY);
            }

            private void DrawStrokeText(string text, double anchorX, double centerY, double height, Color color, bool bold, bool italic, VisioTextHorizontalAlignment? alignment, double rotationRadians, double rotationCenterX, double rotationCenterY) {
                if (string.IsNullOrEmpty(text) || color.A == 0) return;
                double cell = Math.Max(1D, height / 7D);
                double gap = cell * 0.9D;
                double width = MeasureStrokeText(text, height);
                double x = anchorX;
                if (alignment == VisioTextHorizontalAlignment.Right) {
                    x -= width;
                } else if (alignment != VisioTextHorizontalAlignment.Left) {
                    x -= width / 2D;
                }

                double top = centerY - height / 2D;
                double bottom = top + Math.Max(1D, height);
                foreach (char c in text) {
                    DrawGlyph(c, x, top, cell, color, bold, italic, bottom, rotationRadians, rotationCenterX, rotationCenterY);
                    x += (GlyphWidth(c) * cell) + gap;
                }
            }

            private static IReadOnlyList<List<OfficePoint>> TransformContours(IReadOnlyList<List<OfficePoint>> contours, double bottom, bool italic, double rotationRadians, double rotationCenterX, double rotationCenterY) {
                if ((!italic && Math.Abs(rotationRadians) < TextRotationEpsilon) || contours.Count == 0) return contours;
                List<List<OfficePoint>> transformed = new(contours.Count);
                foreach (List<OfficePoint> contour in contours) {
                    List<OfficePoint> points = new(contour.Count);
                    foreach (OfficePoint point in contour) {
                        points.Add(TransformTextPoint(point, bottom, italic, rotationRadians, rotationCenterX, rotationCenterY));
                    }

                    transformed.Add(points);
                }

                return transformed;
            }

            private void FillContours(IReadOnlyList<List<OfficePoint>> contours, Color color) {
                if (color.A == 0 || contours.Count == 0) return;
                (int minX, int minY, int maxX, int maxY) = BoundsContours(contours, 1D);
                for (int y = minY; y <= maxY; y++) {
                    for (int x = minX; x <= maxX; x++) {
                        int hits = 0;
                        for (int i = 0; i < contours.Count; i++) {
                            if (contours[i].Count >= 3 && ContainsPoint(contours[i], x + 0.5D, y + 0.5D)) {
                                hits++;
                            }
                        }

                        if ((hits & 1) == 1) {
                            BlendPixel(x, y, color);
                        }
                    }
                }
            }

            internal byte[] Resolve() {
                if (Supersampling == 1) {
                    return (byte[])_pixels.Clone();
                }

                byte[] output = new byte[_width * _height * 4];
                int samples = Supersampling * Supersampling;
                for (int y = 0; y < _height; y++) {
                    for (int x = 0; x < _width; x++) {
                        int a = 0;
                        long r = 0, g = 0, b = 0;
                        for (int sy = 0; sy < Supersampling; sy++) {
                            for (int sx = 0; sx < Supersampling; sx++) {
                                int source = (((y * Supersampling) + sy) * _renderWidth + ((x * Supersampling) + sx)) * 4;
                                int sampleAlpha = _pixels[source + 3];
                                r += _pixels[source] * sampleAlpha;
                                g += _pixels[source + 1] * sampleAlpha;
                                b += _pixels[source + 2] * sampleAlpha;
                                a += sampleAlpha;
                            }
                        }

                        int target = (y * _width + x) * 4;
                        if (a > 0) {
                            output[target] = (byte)((r + (a / 2L)) / a);
                            output[target + 1] = (byte)((g + (a / 2L)) / a);
                            output[target + 2] = (byte)((b + (a / 2L)) / a);
                        }

                        output[target + 3] = (byte)(a / samples);
                    }
                }

                return output;
            }

            private void StrokeDashedSegment((double X, double Y) start, (double X, double Y) end, Color color, double width) {
                double length = Distance(start, end);
                if (length <= 0D) return;
                double dash = Math.Max(Supersampling * 6D, width * 3D);
                double gap = Math.Max(Supersampling * 4D, width * 2D);
                double pos = 0D;
                while (pos < length) {
                    double next = Math.Min(length, pos + dash);
                    double t1 = pos / length;
                    double t2 = next / length;
                    StrokeSegment(
                        (start.X + ((end.X - start.X) * t1), start.Y + ((end.Y - start.Y) * t1)),
                        (start.X + ((end.X - start.X) * t2), start.Y + ((end.Y - start.Y) * t2)),
                        color,
                        width);
                    pos = next + gap;
                }
            }

            private void StrokeSegment((double X, double Y) start, (double X, double Y) end, Color color, double width) {
                double half = width / 2D;
                int minX = ClampToInt(Math.Floor(Math.Min(start.X, end.X) - half - 1D), 0, _renderWidth - 1);
                int maxX = ClampToInt(Math.Ceiling(Math.Max(start.X, end.X) + half + 1D), 0, _renderWidth - 1);
                int minY = ClampToInt(Math.Floor(Math.Min(start.Y, end.Y) - half - 1D), 0, _renderHeight - 1);
                int maxY = ClampToInt(Math.Ceiling(Math.Max(start.Y, end.Y) + half + 1D), 0, _renderHeight - 1);
                for (int y = minY; y <= maxY; y++) {
                    for (int x = minX; x <= maxX; x++) {
                        double d = DistanceToSegment(x + 0.5D, y + 0.5D, start, end);
                        if (d <= half) {
                            BlendPixel(x, y, color);
                        }
                    }
                }
            }

            internal void FillRectangle(double x, double y, double width, double height, Color color) {
                FillRect(x, y, width, height, color);
            }

            private void DrawGlyph(char c, double x, double y, double cell, Color color, bool bold, bool italic, double bottom, double rotationRadians, double rotationCenterX, double rotationCenterY) {
                string[] rows = GlyphRows(c);
                double strokeWidth = Math.Max(1D, bold ? cell * 0.38D : cell * 0.26D);
                for (int row = 0; row < rows.Length; row++) {
                    string bits = rows[row];
                    for (int col = 0; col < bits.Length; col++) {
                        if (bits[col] != '1') continue;
                        (double X, double Y) current = TransformTextPoint(GlyphPoint(x, y, cell, col, row), bottom, italic, rotationRadians, rotationCenterX, rotationCenterY);
                        bool connected = false;
                        if (col + 1 < bits.Length && bits[col + 1] == '1') {
                            StrokeSegment(current, TransformTextPoint(GlyphPoint(x, y, cell, col + 1, row), bottom, italic, rotationRadians, rotationCenterX, rotationCenterY), color, strokeWidth);
                            connected = true;
                        }

                        if (row + 1 < rows.Length) {
                            string next = rows[row + 1];
                            if (col < next.Length && next[col] == '1') {
                                StrokeSegment(current, TransformTextPoint(GlyphPoint(x, y, cell, col, row + 1), bottom, italic, rotationRadians, rotationCenterX, rotationCenterY), color, strokeWidth);
                                connected = true;
                            }

                            if (col > 0 && col - 1 < next.Length && next[col - 1] == '1') {
                                StrokeSegment(current, TransformTextPoint(GlyphPoint(x, y, cell, col - 1, row + 1), bottom, italic, rotationRadians, rotationCenterX, rotationCenterY), color, strokeWidth);
                                connected = true;
                            }

                            if (col + 1 < next.Length && next[col + 1] == '1') {
                                StrokeSegment(current, TransformTextPoint(GlyphPoint(x, y, cell, col + 1, row + 1), bottom, italic, rotationRadians, rotationCenterX, rotationCenterY), color, strokeWidth);
                                connected = true;
                            }
                        }

                        if (!connected) {
                            DrawEllipse(current.X, current.Y, strokeWidth / 2D, strokeWidth / 2D, color, Color.Transparent, 0D);
                        }
                    }
                }
            }

            private static (double X, double Y) GlyphPoint(double x, double y, double cell, int col, int row) {
                return (x + ((col + 0.5D) * cell), y + ((row + 0.5D) * cell));
            }

            private const double ItalicShear = 0.22D;

            private static OfficePoint SkewItalic(OfficePoint point, double bottom, bool italic) {
                return italic ? new OfficePoint(point.X + ((bottom - point.Y) * ItalicShear), point.Y) : point;
            }

            private static (double X, double Y) SkewItalic((double X, double Y) point, double bottom, bool italic) {
                return italic ? (point.X + ((bottom - point.Y) * ItalicShear), point.Y) : point;
            }

            private static OfficePoint TransformTextPoint(OfficePoint point, double bottom, bool italic, double rotationRadians, double rotationCenterX, double rotationCenterY) {
                if (!italic && Math.Abs(rotationRadians) < TextRotationEpsilon) return point;
                OfficePoint skewed = SkewItalic(point, bottom, italic);
                (double X, double Y) rotated = RotateTextPoint((skewed.X, skewed.Y), rotationCenterX, rotationCenterY, rotationRadians);
                return new OfficePoint(rotated.X, rotated.Y);
            }

            private static (double X, double Y) TransformTextPoint((double X, double Y) point, double bottom, bool italic, double rotationRadians, double rotationCenterX, double rotationCenterY) {
                if (!italic && Math.Abs(rotationRadians) < TextRotationEpsilon) return point;
                return RotateTextPoint(SkewItalic(point, bottom, italic), rotationCenterX, rotationCenterY, rotationRadians);
            }

            private void FillRect(double x, double y, double width, double height, Color color) {
                int minX = ClampToInt(Math.Floor(x), 0, _renderWidth - 1);
                int maxX = ClampToInt(Math.Ceiling(x + width), 0, _renderWidth - 1);
                int minY = ClampToInt(Math.Floor(y), 0, _renderHeight - 1);
                int maxY = ClampToInt(Math.Ceiling(y + height), 0, _renderHeight - 1);
                for (int py = minY; py <= maxY; py++) {
                    for (int px = minX; px <= maxX; px++) {
                        BlendPixel(px, py, color);
                    }
                }
            }

            private void SetPixel(int x, int y, Color color) {
                int offset = (y * _renderWidth + x) * 4;
                _pixels[offset] = color.R;
                _pixels[offset + 1] = color.G;
                _pixels[offset + 2] = color.B;
                _pixels[offset + 3] = color.A;
            }

            private void BlendPixel(int x, int y, Color color) {
                int offset = (y * _renderWidth + x) * 4;
                int srcA = color.A;
                if (srcA == 255 || _pixels[offset + 3] == 0) {
                    _pixels[offset] = color.R;
                    _pixels[offset + 1] = color.G;
                    _pixels[offset + 2] = color.B;
                    _pixels[offset + 3] = color.A;
                    return;
                }

                int dstA = _pixels[offset + 3];
                int outA = srcA + ((dstA * (255 - srcA)) / 255);
                if (outA == 0) return;
                _pixels[offset] = (byte)(((color.R * srcA) + (_pixels[offset] * dstA * (255 - srcA) / 255)) / outA);
                _pixels[offset + 1] = (byte)(((color.G * srcA) + (_pixels[offset + 1] * dstA * (255 - srcA) / 255)) / outA);
                _pixels[offset + 2] = (byte)(((color.B * srcA) + (_pixels[offset + 2] * dstA * (255 - srcA) / 255)) / outA);
                _pixels[offset + 3] = (byte)outA;
            }

            private (int MinX, int MinY, int MaxX, int MaxY) Bounds(IReadOnlyList<(double X, double Y)> points, double pad) {
                double minX = points[0].X;
                double maxX = points[0].X;
                double minY = points[0].Y;
                double maxY = points[0].Y;
                for (int i = 1; i < points.Count; i++) {
                    minX = Math.Min(minX, points[i].X);
                    maxX = Math.Max(maxX, points[i].X);
                    minY = Math.Min(minY, points[i].Y);
                    maxY = Math.Max(maxY, points[i].Y);
                }

                return (
                    ClampToInt(Math.Floor(minX - pad), 0, _renderWidth - 1),
                    ClampToInt(Math.Floor(minY - pad), 0, _renderHeight - 1),
                    ClampToInt(Math.Ceiling(maxX + pad), 0, _renderWidth - 1),
                    ClampToInt(Math.Ceiling(maxY + pad), 0, _renderHeight - 1));
            }

            private (int MinX, int MinY, int MaxX, int MaxY) BoundsContours(IReadOnlyList<List<OfficePoint>> contours, double pad) {
                double minX = double.PositiveInfinity;
                double maxX = double.NegativeInfinity;
                double minY = double.PositiveInfinity;
                double maxY = double.NegativeInfinity;
                for (int i = 0; i < contours.Count; i++) {
                    for (int j = 0; j < contours[i].Count; j++) {
                        minX = Math.Min(minX, contours[i][j].X);
                        maxX = Math.Max(maxX, contours[i][j].X);
                        minY = Math.Min(minY, contours[i][j].Y);
                        maxY = Math.Max(maxY, contours[i][j].Y);
                    }
                }

                if (double.IsInfinity(minX) || double.IsInfinity(minY)) {
                    return (0, 0, -1, -1);
                }

                return (
                    ClampToInt(Math.Floor(minX - pad), 0, _renderWidth - 1),
                    ClampToInt(Math.Floor(minY - pad), 0, _renderHeight - 1),
                    ClampToInt(Math.Ceiling(maxX + pad), 0, _renderWidth - 1),
                    ClampToInt(Math.Ceiling(maxY + pad), 0, _renderHeight - 1));
            }

            private static bool ContainsPoint(IReadOnlyList<(double X, double Y)> points, double x, double y) {
                bool inside = false;
                for (int i = 0, j = points.Count - 1; i < points.Count; j = i++) {
                    if (((points[i].Y > y) != (points[j].Y > y)) &&
                        (x < (points[j].X - points[i].X) * (y - points[i].Y) / (points[j].Y - points[i].Y) + points[i].X)) {
                        inside = !inside;
                    }
                }

                return inside;
            }

            private static bool ContainsPoint(IReadOnlyList<OfficePoint> points, double x, double y) {
                bool inside = false;
                for (int i = 0, j = points.Count - 1; i < points.Count; j = i++) {
                    if (((points[i].Y > y) != (points[j].Y > y)) &&
                        (x < (points[j].X - points[i].X) * (y - points[i].Y) / (points[j].Y - points[i].Y) + points[i].X)) {
                        inside = !inside;
                    }
                }

                return inside;
            }

            private static double DistanceToSegment(double px, double py, (double X, double Y) a, (double X, double Y) b) {
                double dx = b.X - a.X;
                double dy = b.Y - a.Y;
                double lengthSquared = (dx * dx) + (dy * dy);
                if (lengthSquared <= 0D) {
                    double ax = px - a.X;
                    double ay = py - a.Y;
                    return Math.Sqrt((ax * ax) + (ay * ay));
                }

                double t = ((px - a.X) * dx + (py - a.Y) * dy) / lengthSquared;
                t = t < 0D ? 0D : t > 1D ? 1D : t;
                double x = a.X + (t * dx);
                double y = a.Y + (t * dy);
                double sx = px - x;
                double sy = py - y;
                return Math.Sqrt((sx * sx) + (sy * sy));
            }

            private static int ClampToInt(double value, int min, int max) {
                if (value < min) return min;
                if (value > max) return max;
                return (int)value;
            }

            private static int GlyphWidth(char c) => c == ' ' ? 3 : 5;

            private static string[] GlyphRows(char c) {
                switch (char.ToUpperInvariant(c)) {
                    case 'A': return new[] { "01110", "10001", "10001", "11111", "10001", "10001", "10001" };
                    case 'B': return new[] { "11110", "10001", "10001", "11110", "10001", "10001", "11110" };
                    case 'C': return new[] { "01111", "10000", "10000", "10000", "10000", "10000", "01111" };
                    case 'D': return new[] { "11110", "10001", "10001", "10001", "10001", "10001", "11110" };
                    case 'E': return new[] { "11111", "10000", "10000", "11110", "10000", "10000", "11111" };
                    case 'F': return new[] { "11111", "10000", "10000", "11110", "10000", "10000", "10000" };
                    case 'G': return new[] { "01111", "10000", "10000", "10111", "10001", "10001", "01110" };
                    case 'H': return new[] { "10001", "10001", "10001", "11111", "10001", "10001", "10001" };
                    case 'I': return new[] { "11111", "00100", "00100", "00100", "00100", "00100", "11111" };
                    case 'J': return new[] { "00111", "00010", "00010", "00010", "10010", "10010", "01100" };
                    case 'K': return new[] { "10001", "10010", "10100", "11000", "10100", "10010", "10001" };
                    case 'L': return new[] { "10000", "10000", "10000", "10000", "10000", "10000", "11111" };
                    case 'M': return new[] { "10001", "11011", "10101", "10101", "10001", "10001", "10001" };
                    case 'N': return new[] { "10001", "11001", "10101", "10011", "10001", "10001", "10001" };
                    case 'O': return new[] { "01110", "10001", "10001", "10001", "10001", "10001", "01110" };
                    case 'P': return new[] { "11110", "10001", "10001", "11110", "10000", "10000", "10000" };
                    case 'Q': return new[] { "01110", "10001", "10001", "10001", "10101", "10010", "01101" };
                    case 'R': return new[] { "11110", "10001", "10001", "11110", "10100", "10010", "10001" };
                    case 'S': return new[] { "01111", "10000", "10000", "01110", "00001", "00001", "11110" };
                    case 'T': return new[] { "11111", "00100", "00100", "00100", "00100", "00100", "00100" };
                    case 'U': return new[] { "10001", "10001", "10001", "10001", "10001", "10001", "01110" };
                    case 'V': return new[] { "10001", "10001", "10001", "10001", "10001", "01010", "00100" };
                    case 'W': return new[] { "10001", "10001", "10001", "10101", "10101", "10101", "01010" };
                    case 'X': return new[] { "10001", "10001", "01010", "00100", "01010", "10001", "10001" };
                    case 'Y': return new[] { "10001", "10001", "01010", "00100", "00100", "00100", "00100" };
                    case 'Z': return new[] { "11111", "00001", "00010", "00100", "01000", "10000", "11111" };
                    case '0': return new[] { "01110", "10001", "10011", "10101", "11001", "10001", "01110" };
                    case '1': return new[] { "00100", "01100", "00100", "00100", "00100", "00100", "01110" };
                    case '2': return new[] { "01110", "10001", "00001", "00010", "00100", "01000", "11111" };
                    case '3': return new[] { "11110", "00001", "00001", "01110", "00001", "00001", "11110" };
                    case '4': return new[] { "00010", "00110", "01010", "10010", "11111", "00010", "00010" };
                    case '5': return new[] { "11111", "10000", "10000", "11110", "00001", "00001", "11110" };
                    case '6': return new[] { "01110", "10000", "10000", "11110", "10001", "10001", "01110" };
                    case '7': return new[] { "11111", "00001", "00010", "00100", "01000", "01000", "01000" };
                    case '8': return new[] { "01110", "10001", "10001", "01110", "10001", "10001", "01110" };
                    case '9': return new[] { "01110", "10001", "10001", "01111", "00001", "00001", "01110" };
                    case '-': return new[] { "00000", "00000", "00000", "11111", "00000", "00000", "00000" };
                    case '_': return new[] { "00000", "00000", "00000", "00000", "00000", "00000", "11111" };
                    case '+': return new[] { "00000", "00100", "00100", "11111", "00100", "00100", "00000" };
                    case '=': return new[] { "00000", "00000", "11111", "00000", "11111", "00000", "00000" };
                    case '/': return new[] { "00001", "00001", "00010", "00100", "01000", "10000", "10000" };
                    case '\\': return new[] { "10000", "10000", "01000", "00100", "00010", "00001", "00001" };
                    case '.': return new[] { "00000", "00000", "00000", "00000", "00000", "01100", "01100" };
                    case ',': return new[] { "00000", "00000", "00000", "00000", "00000", "01100", "01000" };
                    case ':': return new[] { "00000", "01100", "01100", "00000", "01100", "01100", "00000" };
                    case ';': return new[] { "00000", "01100", "01100", "00000", "01100", "01000", "10000" };
                    case '!': return new[] { "00100", "00100", "00100", "00100", "00100", "00000", "00100" };
                    case '?': return new[] { "01110", "10001", "00001", "00010", "00100", "00000", "00100" };
                    case '&': return new[] { "01100", "10010", "10100", "01000", "10101", "10010", "01101" };
                    case '%': return new[] { "11001", "11010", "00010", "00100", "01000", "01011", "10011" };
                    case '#': return new[] { "01010", "01010", "11111", "01010", "11111", "01010", "01010" };
                    case '(': return new[] { "00010", "00100", "01000", "01000", "01000", "00100", "00010" };
                    case ')': return new[] { "01000", "00100", "00010", "00010", "00010", "00100", "01000" };
                    case '[': return new[] { "01110", "01000", "01000", "01000", "01000", "01000", "01110" };
                    case ']': return new[] { "01110", "00010", "00010", "00010", "00010", "00010", "01110" };
                    case '<': return new[] { "00010", "00100", "01000", "10000", "01000", "00100", "00010" };
                    case '>': return new[] { "01000", "00100", "00010", "00001", "00010", "00100", "01000" };
                    case '|': return new[] { "00100", "00100", "00100", "00100", "00100", "00100", "00100" };
                    case '\'': return new[] { "01100", "00100", "01000", "00000", "00000", "00000", "00000" };
                    case '"': return new[] { "01010", "01010", "01010", "00000", "00000", "00000", "00000" };
                    case ' ': return new[] { "000", "000", "000", "000", "000", "000", "000" };
                    default: return new[] { "11111", "10001", "00001", "00010", "00100", "00000", "00100" };
                }
            }
        }
    }
}
