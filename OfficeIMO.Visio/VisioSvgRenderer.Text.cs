using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Xml;
using Color = OfficeIMO.Drawing.OfficeColor;


namespace OfficeIMO.Visio {
    internal static partial class VisioSvgRenderer {
        private static void WriteShapeText(XmlWriter writer, VisioPage page, VisioShape shape, double scale) {
            VisioTextStyle? style = shape.TextStyle;
            double textWidth = Math.Max(0.05D, style?.TextWidth ?? shape.Width);
            double textHeight = Math.Max(0.05D, style?.TextHeight ?? shape.Height);
            double pinX = style?.TextPinX ?? shape.Width / 2D;
            double pinY = style?.TextPinY ?? shape.Height / 2D;
            (double localX, double localY) = ResolveTextBoxCenter(pinX, pinY, textWidth, textHeight, style);
            (double textX, double textY) = GetPagePoint(shape, localX, localY);
            (double x, double y) = ToSvg(page, textX, textY, scale);
            double horizontalMargins = (style?.LeftMargin ?? 0.05D) + (style?.RightMargin ?? 0.05D);
            double verticalMargins = (style?.TopMargin ?? 0.03D) + (style?.BottomMargin ?? 0.03D);
            WriteText(
                writer,
                shape.Text!,
                x,
                y,
                style,
                defaultSize: 10D,
                scale: scale,
                rotateRadians: shape.Angle + (style?.TextAngle ?? 0D),
                maxWidth: Math.Max(12D, (textWidth - horizontalMargins) * scale),
                maxHeight: Math.Max(8D, (textHeight - verticalMargins) * scale),
                drawLabelBackground: false);
        }

        private static (double X, double Y) ResolveTextBoxCenter(double pinX, double pinY, double width, double height, VisioTextStyle? style) {
            double locPinX = style?.TextLocPinX ?? width / 2D;
            double locPinY = style?.TextLocPinY ?? height / 2D;
            return (pinX + (width / 2D) - locPinX, pinY + (height / 2D) - locPinY);
        }

        private static bool HasVisibleLine(VisioShape shape) =>
            shape.LinePattern != 0 && shape.LineWeight > 0D && shape.LineColor.A > 0;
        private static void WriteText(
            XmlWriter writer,
            string text,
            double x,
            double y,
            VisioTextStyle? style,
            double defaultSize,
            double scale,
            double rotateRadians,
            double maxWidth = 0D,
            double maxHeight = 0D,
            bool drawLabelBackground = false,
            bool labelAdjusted = false) {
            double fontSize = PointsToSvgPixels(style?.Size ?? defaultSize, scale);
            double availableWidth = IsFinitePositive(maxWidth) ? maxWidth : double.PositiveInfinity;
            double availableHeight = IsFinitePositive(maxHeight) ? maxHeight : double.PositiveInfinity;
            TextLayout layout = CreateTextLayout(text, fontSize, availableWidth, availableHeight);
            fontSize = layout.FontSize;
            double anchorX = ResolveTextAnchorX(x, availableWidth, style?.HorizontalAlignment);
            double top = ResolveTextTop(y, layout.Height, availableHeight, style?.VerticalAlignment);

            Color? backgroundColor = ResolveTextBackground(style, drawLabelBackground);
            if (backgroundColor.HasValue) {
                double padX = Math.Max(3D, fontSize * 0.22D);
                double padY = Math.Max(2D, fontSize * 0.16D);
                double backgroundLeft = GetAlignedTextLeft(anchorX, layout.Width, style?.HorizontalAlignment) - padX;
                writer.WriteStartElement("rect", SvgNamespace);
                writer.WriteAttributeString("data-officeimo-text-background", "true");
                if (drawLabelBackground) {
                    writer.WriteAttributeString("data-officeimo-connector-label-background", "true");
                }

                if (labelAdjusted) {
                    writer.WriteAttributeString("data-officeimo-label-adjusted", "true");
                }

                writer.WriteAttributeString("x", Format(backgroundLeft));
                writer.WriteAttributeString("y", Format(top - padY));
                writer.WriteAttributeString("width", Format(layout.Width + (padX * 2D)));
                writer.WriteAttributeString("height", Format(layout.Height + (padY * 2D)));
                if (Math.Abs(rotateRadians) > 1e-9) {
                    writer.WriteAttributeString("transform", FormatTextRotation(rotateRadians, x, y));
                }

                WriteColor(writer, "fill", backgroundColor.Value);
                writer.WriteEndElement();
            }

            writer.WriteStartElement("text", SvgNamespace);
            if (labelAdjusted) {
                writer.WriteAttributeString("data-officeimo-label-adjusted", "true");
            }

            writer.WriteAttributeString("x", Format(anchorX));
            writer.WriteAttributeString("y", Format(top + (fontSize / 2D)));
            writer.WriteAttributeString("font-family", string.IsNullOrWhiteSpace(style?.FontFamily) ? "Aptos, Calibri, Arial, sans-serif" : style!.FontFamily);
            writer.WriteAttributeString("font-size", Format(fontSize));
            writer.WriteAttributeString("text-anchor", GetTextAnchor(style));
            writer.WriteAttributeString("dominant-baseline", "middle");
            WriteColor(writer, "fill", style?.Color ?? Color.FromRgb(17, 24, 39));
            if (style?.Bold == true) writer.WriteAttributeString("font-weight", "700");
            if (style?.Italic == true) writer.WriteAttributeString("font-style", "italic");
            if (style?.Underline == true) writer.WriteAttributeString("text-decoration", "underline");
            if (Math.Abs(rotateRadians) > 1e-9) {
                writer.WriteAttributeString("transform", FormatTextRotation(rotateRadians, x, y));
            }

            for (int i = 0; i < layout.Lines.Length; i++) {
                writer.WriteStartElement("tspan", SvgNamespace);
                writer.WriteAttributeString("x", Format(anchorX));
                writer.WriteAttributeString("dy", i == 0 ? "0" : Format(layout.LineHeight));
                writer.WriteString(layout.Lines[i]);
                writer.WriteEndElement();
            }

            writer.WriteEndElement();
        }

        private static void WriteArrow(
            XmlWriter writer,
            VisioPage page,
            (double X, double Y) tip,
            (double X, double Y) from,
            double scale,
            Color color,
            double strokeWidth,
            string position) {
            (double tipX, double tipY) = ToSvg(page, tip.X, tip.Y, scale);
            (double fromX, double fromY) = ToSvg(page, from.X, from.Y, scale);
            double angle = Math.Atan2(tipY - fromY, tipX - fromX);
            double length = Math.Max(strokeWidth * 4D, 8D);
            double wing = Math.PI / 7D;
            double x1 = tipX - (Math.Cos(angle - wing) * length);
            double y1 = tipY - (Math.Sin(angle - wing) * length);
            double x2 = tipX - (Math.Cos(angle + wing) * length);
            double y2 = tipY - (Math.Sin(angle + wing) * length);

            writer.WriteStartElement("path", SvgNamespace);
            writer.WriteAttributeString("data-officeimo-connector-arrow", position);
            writer.WriteAttributeString("d", "M " + Format(tipX) + " " + Format(tipY) +
                                             " L " + Format(x1) + " " + Format(y1) +
                                             " L " + Format(x2) + " " + Format(y2) + " Z");
            WriteColor(writer, "fill", color);
            writer.WriteAttributeString("stroke", "none");
            writer.WriteEndElement();
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

        private static TextLayout CreateTextLayout(string text, double fontSize, double maxWidth, double maxHeight) {
            string[] lines = WrapText(text, fontSize, maxWidth);
            double lineHeight = fontSize * 1.2D;
            double measuredWidth = MeasureMaxLineWidth(lines, fontSize);
            double measuredHeight = Math.Max(fontSize, ((lines.Length - 1) * lineHeight) + fontSize);
            double scaleDown = Math.Min(1D, Math.Min(maxWidth / Math.Max(measuredWidth, 1D), maxHeight / Math.Max(measuredHeight, 1D)));
            if (scaleDown < 0.98D) {
                fontSize = Math.Max(5D, fontSize * scaleDown);
                lines = WrapText(text, fontSize, maxWidth);
                lineHeight = fontSize * 1.2D;
                measuredWidth = MeasureMaxLineWidth(lines, fontSize);
                measuredHeight = Math.Max(fontSize, ((lines.Length - 1) * lineHeight) + fontSize);
            }

            return new TextLayout(lines, fontSize, lineHeight, measuredWidth, measuredHeight);
        }

        private static string[] WrapText(string text, double fontSize, double maxWidth) {
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
                    if (EstimateTextWidth(word, fontSize) > maxWidth) {
                        if (current.Length > 0) {
                            output.Add(current);
                            current = string.Empty;
                        }

                        foreach (string part in BreakWord(word, fontSize, maxWidth)) {
                            output.Add(part);
                        }

                        continue;
                    }

                    string candidate = current.Length == 0 ? word : current + " " + word;
                    if (current.Length > 0 && EstimateTextWidth(candidate, fontSize) > maxWidth) {
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

        private static IEnumerable<string> BreakWord(string word, double fontSize, double maxWidth) {
            StringBuilder part = new();
            foreach (char c in word) {
                string candidate = part.ToString() + c;
                if (part.Length > 0 && EstimateTextWidth(candidate, fontSize) > maxWidth) {
                    yield return part.ToString();
                    part.Clear();
                }

                part.Append(c);
            }

            if (part.Length > 0) {
                yield return part.ToString();
            }
        }

        private static double MeasureMaxLineWidth(IReadOnlyList<string> lines, double fontSize) {
            double max = 0D;
            for (int i = 0; i < lines.Count; i++) {
                max = Math.Max(max, EstimateTextWidth(lines[i], fontSize));
            }

            return max;
        }

        private static double EstimateTextWidth(string text, double fontSize) {
            double width = 0D;
            foreach (char c in text) {
                if (char.IsWhiteSpace(c)) {
                    width += fontSize * 0.32D;
                } else if ("ilI.,'!:;|".IndexOf(c) >= 0) {
                    width += fontSize * 0.26D;
                } else if ("MW@#%&".IndexOf(c) >= 0) {
                    width += fontSize * 0.86D;
                } else if (char.IsDigit(c)) {
                    width += fontSize * 0.56D;
                } else {
                    width += fontSize * 0.54D;
                }
            }

            return width;
        }

        private static double ResolveTextAnchorX(double centerX, double maxWidth, VisioTextHorizontalAlignment? alignment) {
            if (!IsFinitePositive(maxWidth)) {
                return centerX;
            }

            switch (alignment) {
                case VisioTextHorizontalAlignment.Left:
                    return centerX - (maxWidth / 2D);
                case VisioTextHorizontalAlignment.Right:
                    return centerX + (maxWidth / 2D);
                default:
                    return centerX;
            }
        }

        private static double ResolveTextTop(double centerY, double measuredHeight, double maxHeight, VisioTextVerticalAlignment? alignment) {
            if (!IsFinitePositive(maxHeight)) {
                return centerY - (measuredHeight / 2D);
            }

            switch (alignment) {
                case VisioTextVerticalAlignment.Top:
                    return centerY - (maxHeight / 2D);
                case VisioTextVerticalAlignment.Bottom:
                    return centerY + (maxHeight / 2D) - measuredHeight;
                default:
                    return centerY - (measuredHeight / 2D);
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

        private static bool IsFinitePositive(double value) =>
            value > 0D && !double.IsNaN(value) && !double.IsInfinity(value);

        private static Color? ResolveTextBackground(VisioTextStyle? style, bool drawLabelBackground) {
            if (style?.BackgroundColor.HasValue == true) {
                return ApplyBackgroundTransparency(style.BackgroundColor.Value, style.BackgroundTransparency);
            }

            return drawLabelBackground ? Color.FromRgba(255, 255, 255, 230) : null;
        }

        private static Color ApplyBackgroundTransparency(Color color, double? transparency) {
            if (!transparency.HasValue) {
                return color;
            }

            double clamped = Math.Max(0D, Math.Min(100D, transparency.Value));
            byte alpha = (byte)Math.Round(color.A * (1D - (clamped / 100D)));
            return Color.FromRgba(color.R, color.G, color.B, alpha);
        }

        private static string FormatTextRotation(double radians, double centerX, double centerY) =>
            "rotate(" + Format(RadiansToDegrees(-radians)) + " " + Format(centerX) + " " + Format(centerY) + ")";

        private static string GetTextAnchor(VisioTextStyle? style) {
            switch (style?.HorizontalAlignment) {
                case VisioTextHorizontalAlignment.Left:
                    return "start";
                case VisioTextHorizontalAlignment.Right:
                    return "end";
                default:
                    return "middle";
            }
        }

        private sealed class TextLayout {
            internal TextLayout(string[] lines, double fontSize, double lineHeight, double width, double height) {
                Lines = lines;
                FontSize = fontSize;
                LineHeight = lineHeight;
                Width = width;
                Height = height;
            }

            internal string[] Lines { get; }

            internal double FontSize { get; }

            internal double LineHeight { get; }

            internal double Width { get; }

            internal double Height { get; }
        }
    }
}
