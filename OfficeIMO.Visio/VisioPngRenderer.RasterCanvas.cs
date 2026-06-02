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

            internal void FillPolygonsEvenOdd(IReadOnlyList<List<(double X, double Y)>> contours, Color color) {
                if (color.A == 0 || contours.Count == 0) return;
                (int minX, int minY, int maxX, int maxY) = BoundsPolygons(contours, 1D);
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

            private (int MinX, int MinY, int MaxX, int MaxY) BoundsPolygons(IReadOnlyList<List<(double X, double Y)>> contours, double pad) {
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
