using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio {
    internal static partial class VisioSvgPreviewRasterizer {
        private static bool RenderText(OfficeRasterCanvas canvas, XElement element, SvgPaint paint, SvgTransform transform, SvgRenderContext context) {
            SvgTextStyle style = SvgTextStyle.Resolve(element, SvgTextStyle.Default, context);
            double x = ReadLength(element, "x", 0D);
            double y = ReadLength(element, "y", style.FontSize);
            double cursorX = x + ReadLength(element, "dx", 0D);
            double cursorY = y + ReadLength(element, "dy", 0D);
            bool rendered = false;

            foreach (XNode node in element.Nodes()) {
                if (node is XText textNode) {
                    rendered |= DrawSvgTextRun(canvas, NormalizeText(textNode.Value), cursorX, cursorY, paint, style, transform, out double advance);
                    cursorX += advance;
                    continue;
                }

                if (node is XElement child && string.Equals(child.Name.LocalName, "tspan", StringComparison.OrdinalIgnoreCase)) {
                    SvgTextStyle childStyle = SvgTextStyle.Resolve(child, style, context);
                    double childX = ReadLength(child, "x", cursorX);
                    double childY = ReadLength(child, "y", cursorY);
                    childX += ReadLength(child, "dx", 0D);
                    childY += ReadLength(child, "dy", 0D);
                    SvgPaint childPaint = SvgPaint.Resolve(child, paint, context);
                    string value = NormalizeText(child.Value);
                    rendered |= DrawSvgTextRun(canvas, value, childX, childY, childPaint, childStyle, transform, out double advance);
                    cursorX = childX + advance;
                    cursorY = childY;
                }
            }

            return rendered;
        }

        private static bool DrawSvgTextRun(
            OfficeRasterCanvas canvas,
            string text,
            double x,
            double baselineY,
            SvgPaint paint,
            SvgTextStyle style,
            SvgTransform transform,
            out double advance) {
            advance = 0D;
            if (string.IsNullOrEmpty(text) || paint.Fill.A == 0 || style.FontSize <= 0D) {
                return false;
            }

            OfficePoint anchor = transform.Apply(x, baselineY);
            double fontHeight = Math.Max(1D, style.FontSize * transform.StrokeScale);
            double top = anchor.Y - (fontHeight * style.BaselineOffset);
            canvas.DrawTextLine(
                text,
                anchor.X,
                top,
                fontHeight,
                paint.Fill,
                bold: style.Bold,
                italic: style.Italic,
                alignment: style.Alignment,
                rotationDegrees: transform.RotationDegrees,
                rotationCenterX: anchor.X,
                rotationCenterY: anchor.Y,
                underline: style.Underline,
                strikethrough: style.Strikethrough,
                fontFamily: style.FontFamily);
            advance = canvas.MeasureText(text, fontHeight, style.FontFamily) / Math.Max(0.0001D, transform.StrokeScale);
            return true;
        }

        private static string NormalizeText(string? text) {
            if (string.IsNullOrWhiteSpace(text)) {
                return string.Empty;
            }

            StringBuilder builder = new(text!.Length);
            bool pendingSpace = false;
            for (int i = 0; i < text.Length; i++) {
                if (char.IsWhiteSpace(text[i])) {
                    pendingSpace = builder.Length > 0;
                    continue;
                }

                if (pendingSpace) {
                    builder.Append(' ');
                    pendingSpace = false;
                }

                builder.Append(text[i]);
            }

            return builder.ToString();
        }

        private readonly struct SvgTextStyle {
            internal static SvgTextStyle Default => new(12D, null, false, false, false, false, OfficeTextAlignment.Left, 0.8D);

            private SvgTextStyle(
                double fontSize,
                string? fontFamily,
                bool bold,
                bool italic,
                bool underline,
                bool strikethrough,
                OfficeTextAlignment alignment,
                double baselineOffset) {
                FontSize = fontSize;
                FontFamily = fontFamily;
                Bold = bold;
                Italic = italic;
                Underline = underline;
                Strikethrough = strikethrough;
                Alignment = alignment;
                BaselineOffset = baselineOffset;
            }

            internal double FontSize { get; }

            internal string? FontFamily { get; }

            internal bool Bold { get; }

            internal bool Italic { get; }

            internal bool Underline { get; }

            internal bool Strikethrough { get; }

            internal OfficeTextAlignment Alignment { get; }

            internal double BaselineOffset { get; }

            internal static SvgTextStyle Resolve(XElement element, SvgTextStyle inherited, SvgRenderContext context) {
                Dictionary<string, string> style = context.StyleSheet.CreateStyle(element);
                double fontSize = ReadStyleLength(element, style, "font-size", inherited.FontSize);
                string? fontFamily = ReadStyleString(element, style, "font-family") ?? inherited.FontFamily;
                bool bold = ReadFontWeight(element, style, inherited.Bold);
                bool italic = ReadFontStyle(element, style, inherited.Italic);
                bool underline = inherited.Underline;
                bool strikethrough = inherited.Strikethrough;
                ReadTextDecoration(element, style, ref underline, ref strikethrough);
                OfficeTextAlignment alignment = ReadTextAnchor(element, style, inherited.Alignment);
                double baselineOffset = ReadBaselineOffset(element, style, inherited.BaselineOffset);
                return new SvgTextStyle(fontSize, fontFamily, bold, italic, underline, strikethrough, alignment, baselineOffset);
            }

            private static double ReadStyleLength(XElement element, Dictionary<string, string> style, string name, double fallback) {
                string? raw = element.Attribute(name)?.Value ?? (style.TryGetValue(name, out string? value) ? value : null);
                return TryParseLength(raw, out double parsed) ? parsed : fallback;
            }

            private static string? ReadStyleString(XElement element, Dictionary<string, string> style, string name) {
                string? raw = element.Attribute(name)?.Value ?? (style.TryGetValue(name, out string? value) ? value : null);
                if (string.IsNullOrWhiteSpace(raw)) {
                    return null;
                }

                return raw!.Trim().Trim('\'', '"');
            }

            private static bool ReadFontWeight(XElement element, Dictionary<string, string> style, bool inherited) {
                string? raw = ReadStyleString(element, style, "font-weight");
                if (string.IsNullOrWhiteSpace(raw)) {
                    return inherited;
                }

                if (string.Equals(raw, "normal", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(raw, "400", StringComparison.OrdinalIgnoreCase)) {
                    return false;
                }

                if (int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out int numeric)) {
                    return numeric >= 600;
                }

                return string.Equals(raw, "bold", StringComparison.OrdinalIgnoreCase) ||
                       string.Equals(raw, "bolder", StringComparison.OrdinalIgnoreCase);
            }

            private static bool ReadFontStyle(XElement element, Dictionary<string, string> style, bool inherited) {
                string? raw = ReadStyleString(element, style, "font-style");
                if (string.IsNullOrWhiteSpace(raw)) {
                    return inherited;
                }

                if (string.Equals(raw, "normal", StringComparison.OrdinalIgnoreCase)) {
                    return false;
                }

                return string.Equals(raw, "italic", StringComparison.OrdinalIgnoreCase) ||
                       string.Equals(raw, "oblique", StringComparison.OrdinalIgnoreCase);
            }

            private static void ReadTextDecoration(XElement element, Dictionary<string, string> style, ref bool underline, ref bool strikethrough) {
                string? raw = ReadStyleString(element, style, "text-decoration") ?? ReadStyleString(element, style, "text-decoration-line");
                if (string.IsNullOrWhiteSpace(raw)) {
                    return;
                }

                if (string.Equals(raw, "none", StringComparison.OrdinalIgnoreCase)) {
                    underline = false;
                    strikethrough = false;
                    return;
                }

                string[] parts = raw!.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
                for (int i = 0; i < parts.Length; i++) {
                    if (string.Equals(parts[i], "underline", StringComparison.OrdinalIgnoreCase)) {
                        underline = true;
                    } else if (string.Equals(parts[i], "line-through", StringComparison.OrdinalIgnoreCase)) {
                        strikethrough = true;
                    }
                }
            }

            private static OfficeTextAlignment ReadTextAnchor(XElement element, Dictionary<string, string> style, OfficeTextAlignment inherited) {
                string? raw = ReadStyleString(element, style, "text-anchor");
                if (string.IsNullOrWhiteSpace(raw)) {
                    return inherited;
                }

                if (string.Equals(raw, "middle", StringComparison.OrdinalIgnoreCase)) {
                    return OfficeTextAlignment.Center;
                }

                if (string.Equals(raw, "end", StringComparison.OrdinalIgnoreCase)) {
                    return OfficeTextAlignment.Right;
                }

                return OfficeTextAlignment.Left;
            }

            private static double ReadBaselineOffset(XElement element, Dictionary<string, string> style, double inherited) {
                string? raw = ReadStyleString(element, style, "dominant-baseline") ?? ReadStyleString(element, style, "alignment-baseline");
                if (string.IsNullOrWhiteSpace(raw)) {
                    return inherited;
                }

                if (string.Equals(raw, "middle", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(raw, "central", StringComparison.OrdinalIgnoreCase)) {
                    return 0.45D;
                }

                if (string.Equals(raw, "hanging", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(raw, "text-before-edge", StringComparison.OrdinalIgnoreCase)) {
                    return 0D;
                }

                return 0.8D;
            }
        }
    }
}
