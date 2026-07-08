using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio {
    internal static partial class VisioSvgPreviewRasterizer {
        private static bool RenderText(OfficeRasterCanvas canvas, XElement element, SvgPaint paint, SvgTransform transform, SvgRenderContext context) {
            SvgTextStyle style = context.CurrentTextStyle;
            double x = ReadLength(element, "x", 0D, context, SvgLengthAxis.X);
            double y = ReadLength(element, "y", style.FontSize, context, SvgLengthAxis.Y);
            var cursor = new SvgTextCursor(
                x + ReadLength(element, "dx", 0D, context, SvgLengthAxis.X),
                y + ReadLength(element, "dy", 0D, context, SvgLengthAxis.Y));
            cursor.X = ApplyTextAnchor(cursor.X, MeasureTextChunk(canvas, element, style, context, stopAtPositionedChild: true), style.Alignment);
            return RenderTextNodes(canvas, element, paint, style, transform, context, ref cursor);
        }

        private static bool RenderTextNodes(
            OfficeRasterCanvas canvas,
            XElement element,
            SvgPaint paint,
            SvgTextStyle style,
            SvgTransform transform,
            SvgRenderContext context,
            ref SvgTextCursor cursor) {
            bool rendered = false;

            foreach (XNode node in element.Nodes()) {
                if (node is XText textNode) {
                    if (!context.IsVisible) {
                        continue;
                    }

                    string value = NormalizeTextRun(textNode.Value, style.PreserveWhitespace, ref cursor.PendingSpace, cursor.HasTextRun);
                    rendered |= DrawSvgTextRun(canvas, value, cursor.X, cursor.Y, paint, style, transform, out double advance);
                    cursor.X += advance;
                    cursor.HasTextRun |= value.Length > 0;
                    continue;
                }

                if (node is XElement child && string.Equals(child.Name.LocalName, "tspan", StringComparison.OrdinalIgnoreCase)) {
                    if (IsElementDisplayNone(child, context)) {
                        continue;
                    }

                    using IDisposable visibilityScope = context.PushVisibility(ReadVisibilityOverride(child, context));
                    if (!context.IsVisible) {
                        continue;
                    }

                    bool resetsTextFlow = child.Attribute("x") != null || child.Attribute("y") != null;
                    if (resetsTextFlow) {
                        cursor.PendingSpace = false;
                    }

                    SvgTextStyle childStyle = SvgTextStyle.Resolve(child, style, context);
                    double childX = ReadLength(child, "x", cursor.X, context, SvgLengthAxis.X);
                    double childY = ReadLength(child, "y", cursor.Y, context, SvgLengthAxis.Y);
                    childX += ReadLength(child, "dx", 0D, context, SvgLengthAxis.X);
                    childY += ReadLength(child, "dy", 0D, context, SvgLengthAxis.Y);
                    if (resetsTextFlow) {
                        childX = ApplyTextAnchor(childX, MeasureTextChunk(canvas, child, childStyle, context, stopAtPositionedChild: true), childStyle.Alignment);
                    }

                    SvgPaint childPaint = SvgPaint.Resolve(child, paint, context);
                    var childCursor = new SvgTextCursor(childX, childY) {
                        HasTextRun = cursor.HasTextRun && !resetsTextFlow,
                        PendingSpace = cursor.PendingSpace
                    };
                    rendered |= RenderTextNodes(canvas, child, childPaint, childStyle, transform, context, ref childCursor);
                    cursor.X = childCursor.X;
                    cursor.Y = childCursor.Y;
                    cursor.PendingSpace = childCursor.PendingSpace;
                    cursor.HasTextRun |= childCursor.HasTextRun;
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
            OfficeColor textColor = paint.Fill.A > 0 ? paint.Fill : GetStrokeFallbackColor(paint);
            if (string.IsNullOrEmpty(text) || textColor.A == 0 || style.FontSize <= 0D) {
                return false;
            }

            double fontHeight = Math.Max(1D, style.FontSize);
            double top = baselineY - (fontHeight * style.BaselineOffset);
            canvas.DrawTextLineTransformed(
                text,
                x,
                top,
                fontHeight,
                textColor,
                transform.ToOfficeTransform(),
                bold: style.Bold,
                italic: style.Italic,
                alignment: OfficeTextAlignment.Left,
                underline: style.Underline,
                strikethrough: style.Strikethrough,
                fontFamily: style.FontFamily);
            advance = canvas.MeasureText(text, fontHeight, style.FontFamily);
            return true;
        }

        private static double MeasureTextChunk(OfficeRasterCanvas canvas, XElement element, SvgTextStyle style, SvgRenderContext context, bool stopAtPositionedChild) {
            double width = 0D;
            var measureCursor = new SvgTextCursor(0D, 0D);
            foreach (XNode node in element.Nodes()) {
                if (node is XText textNode) {
                    if (!context.IsVisible) {
                        continue;
                    }

                    string value = NormalizeTextRun(textNode.Value, style.PreserveWhitespace, ref measureCursor.PendingSpace, measureCursor.HasTextRun);
                    if (value.Length > 0) {
                        width += canvas.MeasureText(value, Math.Max(1D, style.FontSize), style.FontFamily);
                        measureCursor.HasTextRun = true;
                    }

                    continue;
                }

                if (node is XElement child && string.Equals(child.Name.LocalName, "tspan", StringComparison.OrdinalIgnoreCase)) {
                    if (IsElementDisplayNone(child, context)) {
                        continue;
                    }

                    using IDisposable visibilityScope = context.PushVisibility(ReadVisibilityOverride(child, context));
                    if (!context.IsVisible) {
                        continue;
                    }

                    bool resetsTextFlow = child.Attribute("x") != null || child.Attribute("y") != null;
                    if (resetsTextFlow && stopAtPositionedChild) {
                        break;
                    }

                    SvgTextStyle childStyle = SvgTextStyle.Resolve(child, style, context);
                    width += MeasureTextChunk(canvas, child, childStyle, context, stopAtPositionedChild: true, measureCursor.PendingSpace, measureCursor.HasTextRun);
                    measureCursor.PendingSpace = false;
                    measureCursor.HasTextRun = true;
                }
            }

            return width;
        }

        private static double MeasureTextChunk(OfficeRasterCanvas canvas, XElement element, SvgTextStyle style, SvgRenderContext context, bool stopAtPositionedChild, bool pendingSpace, bool hasPriorTextRun) {
            double width = 0D;
            var measureCursor = new SvgTextCursor(0D, 0D) {
                PendingSpace = pendingSpace,
                HasTextRun = hasPriorTextRun
            };
            foreach (XNode node in element.Nodes()) {
                if (node is XText textNode) {
                    if (!context.IsVisible) {
                        continue;
                    }

                    string value = NormalizeTextRun(textNode.Value, style.PreserveWhitespace, ref measureCursor.PendingSpace, measureCursor.HasTextRun);
                    if (value.Length > 0) {
                        width += canvas.MeasureText(value, Math.Max(1D, style.FontSize), style.FontFamily);
                        measureCursor.HasTextRun = true;
                    }

                    continue;
                }

                if (node is XElement child && string.Equals(child.Name.LocalName, "tspan", StringComparison.OrdinalIgnoreCase)) {
                    if (IsElementDisplayNone(child, context)) {
                        continue;
                    }

                    using IDisposable visibilityScope = context.PushVisibility(ReadVisibilityOverride(child, context));
                    if (!context.IsVisible) {
                        continue;
                    }

                    bool resetsTextFlow = child.Attribute("x") != null || child.Attribute("y") != null;
                    if (resetsTextFlow && stopAtPositionedChild) {
                        break;
                    }

                    SvgTextStyle childStyle = SvgTextStyle.Resolve(child, style, context);
                    width += MeasureTextChunk(canvas, child, childStyle, context, stopAtPositionedChild: true, measureCursor.PendingSpace, measureCursor.HasTextRun);
                    measureCursor.PendingSpace = false;
                    measureCursor.HasTextRun = true;
                }
            }

            return width;
        }

        private static double ApplyTextAnchor(double x, double width, OfficeTextAlignment alignment) {
            if (width <= 0D) {
                return x;
            }

            switch (alignment) {
                case OfficeTextAlignment.Center:
                    return x - (width / 2D);
                case OfficeTextAlignment.Right:
                    return x - width;
                default:
                    return x;
            }
        }

        private static string NormalizeTextRun(string? text, bool preserveWhitespace, ref bool pendingSpace, bool hasPriorTextRun) {
            if (preserveWhitespace) {
                pendingSpace = false;
                return text ?? string.Empty;
            }

            if (string.IsNullOrWhiteSpace(text)) {
                if (hasPriorTextRun && text != null) {
                    for (int i = 0; i < text.Length; i++) {
                        if (char.IsWhiteSpace(text[i])) {
                            pendingSpace = true;
                            break;
                        }
                    }
                }

                return string.Empty;
            }

            StringBuilder builder = new(text!.Length);
            for (int i = 0; i < text.Length; i++) {
                if (char.IsWhiteSpace(text[i])) {
                    pendingSpace = builder.Length > 0 || hasPriorTextRun;
                    continue;
                }

                if (pendingSpace && (builder.Length > 0 || hasPriorTextRun)) {
                    builder.Append(' ');
                }

                pendingSpace = false;
                builder.Append(text[i]);
            }

            return builder.ToString();
        }

        private struct SvgTextCursor {
            internal SvgTextCursor(double x, double y) {
                X = x;
                Y = y;
                PendingSpace = false;
                HasTextRun = false;
            }

            internal double X;

            internal double Y;

            internal bool PendingSpace;

            internal bool HasTextRun;
        }

        private readonly struct SvgTextStyle {
            internal static SvgTextStyle Default => new(12D, null, false, false, false, false, false, OfficeTextAlignment.Left, 0.8D);

            private SvgTextStyle(
                double fontSize,
                string? fontFamily,
                bool bold,
                bool italic,
                bool underline,
                bool strikethrough,
                bool preserveWhitespace,
                OfficeTextAlignment alignment,
                double baselineOffset) {
                FontSize = fontSize;
                FontFamily = fontFamily;
                Bold = bold;
                Italic = italic;
                Underline = underline;
                Strikethrough = strikethrough;
                PreserveWhitespace = preserveWhitespace;
                Alignment = alignment;
                BaselineOffset = baselineOffset;
            }

            internal double FontSize { get; }

            internal string? FontFamily { get; }

            internal bool Bold { get; }

            internal bool Italic { get; }

            internal bool Underline { get; }

            internal bool Strikethrough { get; }

            internal bool PreserveWhitespace { get; }

            internal OfficeTextAlignment Alignment { get; }

            internal double BaselineOffset { get; }

            internal static SvgTextStyle Resolve(XElement element, SvgTextStyle inherited, SvgRenderContext context) {
                Dictionary<string, string> style = context.StyleSheet.CreateStyle(element);
                double fontSize = ReadStyleLength(element, style, "font-size", inherited.FontSize, context);
                string? fontFamily = ReadStyleString(element, style, "font-family") ?? inherited.FontFamily;
                bool bold = ReadFontWeight(element, style, inherited.Bold);
                bool italic = ReadFontStyle(element, style, inherited.Italic);
                bool underline = inherited.Underline;
                bool strikethrough = inherited.Strikethrough;
                ReadTextDecoration(element, style, ref underline, ref strikethrough);
                bool preserveWhitespace = ReadPreserveWhitespace(element, style, inherited.PreserveWhitespace);
                OfficeTextAlignment alignment = ReadTextAnchor(element, style, inherited.Alignment);
                double baselineOffset = ReadBaselineOffset(element, style, inherited.BaselineOffset);
                return new SvgTextStyle(fontSize, fontFamily, bold, italic, underline, strikethrough, preserveWhitespace, alignment, baselineOffset);
            }

            private static double ReadStyleLength(XElement element, Dictionary<string, string> style, string name, double fallback, SvgRenderContext context) {
                string? raw = style.TryGetValue(name, out string? value) ? value : element.Attribute(name)?.Value;
                return TryParseLength(raw, GetLengthReference(context, SvgLengthAxis.Diagonal), out double parsed) ? parsed : fallback;
            }

            private static string? ReadStyleString(XElement element, Dictionary<string, string> style, string name) {
                string? raw = style.TryGetValue(name, out string? value) ? value : element.Attribute(name)?.Value;
                if (string.IsNullOrWhiteSpace(raw)) {
                    return null;
                }

                return raw!.Trim().Trim('\'', '"');
            }

            private static bool ReadPreserveWhitespace(XElement element, Dictionary<string, string> style, bool inherited) {
                string? xmlSpace = element.Attribute(XNamespace.Xml + "space")?.Value;
                if (!string.IsNullOrWhiteSpace(xmlSpace)) {
                    if (string.Equals(xmlSpace, "preserve", StringComparison.OrdinalIgnoreCase)) {
                        return true;
                    }

                    if (string.Equals(xmlSpace, "default", StringComparison.OrdinalIgnoreCase)) {
                        return false;
                    }
                }

                string? whiteSpace = ReadStyleString(element, style, "white-space");
                if (string.IsNullOrWhiteSpace(whiteSpace)) {
                    return inherited;
                }

                return string.Equals(whiteSpace, "pre", StringComparison.OrdinalIgnoreCase) ||
                       string.Equals(whiteSpace, "pre-wrap", StringComparison.OrdinalIgnoreCase) ||
                       string.Equals(whiteSpace, "break-spaces", StringComparison.OrdinalIgnoreCase);
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
