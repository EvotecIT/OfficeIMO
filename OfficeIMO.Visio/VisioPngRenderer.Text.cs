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
    }
}
