using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void DrawDrawingTextAt(OfficeDrawingText text, double originX, double originTopY, OfficeDrawingTextMetrics textMetrics) {
            if (string.IsNullOrEmpty(text.Text)) return;
            if (!text.WrapText && !text.ShrinkToFit && !text.StackedText && !text.HasPadding
                && text.VerticalAlignment == OfficeTextVerticalAlignment.Top) {
                DrawDrawingPositionedText(text, originX, originTopY, textMetrics.MeasureText);
                return;
            }

            string value = text.StackedText ? StackTextElements(text.Text) : text.Text;
            int priorBaselineLevel = text.BaselineLevel == 0
                ? 0
                : text.BaselineLevel - Math.Sign(text.BaselineLevel);
            OfficeTextScriptGeometry priorScript = OfficeTextScriptGeometry.Resolve(
                Math.Max(0.001D, text.Font.Size),
                priorBaselineLevel,
                superscriptOffsetFactor: 0.35D,
                subscriptOffsetFactor: 0.18D);
            var run = new OfficeRichTextRun(
                value,
                priorScript.RenderedFontSize,
                text.Color ?? OfficeColor.Black,
                text.Font.IsBold,
                text.Font.IsItalic,
                text.Font.IsUnderline,
                text.Font.FamilyName,
                text.Font.IsStrikethrough,
                underlineStyle: text.UnderlineStyle,
                strikethroughStyle: text.StrikethroughStyle,
                baseline: text.Baseline);
            var richText = new OfficeDrawingRichText(
                new[] { run },
                text.X,
                text.Y,
                text.Width,
                text.Height,
                text.Alignment,
                text.LineHeight,
                text.VerticalAlignment,
                text.RotationDegrees,
                text.RotationCenterX,
                text.RotationCenterY,
                text.WrapText || text.StackedText,
                text.ShrinkToFit,
                text.FlipHorizontal,
                text.FlipVertical,
                text.Padding,
                text.ParagraphIndent);
            DrawDrawingRichTextAt(richText, originX, originTopY - priorScript.BaselineOffset, textMetrics, text.DecorationColor);
        }

        private void DrawDrawingPositionedText(OfficeDrawingText text, double originX, double originTopY,
            Func<string?, double, string?, OfficeFontStyle, double> measure) {
            double size = text.Font.Size * text.BaselineScale;
            double frameX = originX + text.X;
            double frameTopY = originTopY - text.Y;
            void Paint() {
                double baseline = frameTopY - text.Font.Size - text.BaselineOffset;
                string[] lines = text.Text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
                foreach (string value in lines) {
                    double advance = lines.Length == 1 && text.TextAdvanceWidth.HasValue
                        ? text.TextAdvanceWidth.Value
                        : measure(value, size, text.Font.FamilyName, text.Font.Style);
                    double x = OfficeTextPlacement.ResolveLineLeft(frameX, text.Width, advance, text.Alignment);
                    var run = new PdfTextRun(value, text.Font.IsBold, text.Font.IsUnderline,
                        ToPdfColor(text.Color ?? OfficeColor.Black), text.Font.IsItalic, text.Font.IsStrikethrough,
                        size, ResolveDrawingTextFont(text.Font.FamilyName), fontFamily: text.Font.FamilyName,
                        underlineStyle: text.UnderlineStyle, strikeStyle: text.StrikethroughStyle,
                        decorationColor: ToPdfColor(text.DecorationColor)).WithFeatureSettings(text.FeatureSettings);
                    WriteDrawingPositionedRun(run, x, baseline, advance, frameX, frameTopY - text.Height, text.Width, text.Height);
                    baseline -= text.LineHeight ?? text.Font.Size * 1.2D;
                }
            }
            if (text.HasFrameTransform) {
                OfficeTransform transform = ToTopLeftPageTransform(text.CreateFrameTransform().CreateDestinationTransform(), originX, originTopY);
                RenderEffectGroup(transform, 1D, Paint);
            } else {
                Paint();
            }
        }

        private void DrawDrawingRichTextAt(OfficeDrawingRichText text, double originX, double originTopY, OfficeDrawingTextMetrics textMetrics, OfficeColor? decorationColor = null) {
            if (text.Runs.Count == 0 || string.IsNullOrEmpty(text.PlainText)) return;

            void DrawContent() => DrawDrawingRichTextCore(text, originX + text.X, originTopY - text.Y, textMetrics, decorationColor);
            if (text.HasFrameTransform) {
                OfficeTransform pageTransform = ToTopLeftPageTransform(
                    text.CreateFrameTransform().CreateDestinationTransform(),
                    originX,
                    originTopY);
                RenderEffectGroup(pageTransform, 1D, DrawContent);
            } else {
                DrawContent();
            }
        }

        private void DrawDrawingRichTextCore(OfficeDrawingRichText text, double frameX, double frameTopY,
            OfficeDrawingTextMetrics textMetrics, OfficeColor? decorationColor) {
            double contentX = frameX + text.Padding.Left;
            double contentTopY = frameTopY - text.Padding.Top;
            double width = text.Width - text.Padding.Horizontal;
            double height = text.Height - text.Padding.Vertical;
            if (width <= 0D || height <= 0D) return;

            OfficeRichTextBlockLayout layout = OfficeDrawingTextLayout.Create(text, width, height, textMetrics.MeasureText, measurePaint: textMetrics.MeasurePaintBounds);
            double lineTop = OfficeTextPlacement.ResolveTop(0D, height, layout.Height, text.VerticalAlignment) + layout.ContentOffsetY;
            for (int index = 0; index < layout.Lines.Count; index++) {
                OfficeRichTextLine line = layout.Lines[index];
                double lineHeight = OfficeTextBlockRenderer.ResolveRichTextRenderLineHeight(line, layout.LineHeight);
                double baseline = OfficeTextBlockRenderer.ResolveRichTextRenderBaseline(line, lineTop, lineHeight, true);
                double lineLeft = contentX + line.OffsetX;
                double lineWidth = Math.Max(0D, width - line.OffsetX);
                bool justify = OfficeTextBlockRenderer.ShouldJustifyRichTextLine(line, index, layout.Lines.Count, lineWidth, text.Alignment);
                double cursor = OfficeTextPlacement.ResolveLineLeft(lineLeft, lineWidth, line.Width, text.Alignment);
                if (justify) {
                    var tokens = OfficeTextBlockRenderer.CreateRichTextRenderTokens(line, textMetrics.MeasureText);
                    int gaps = OfficeTextBlockRenderer.CountJustifiableRichTextGaps(tokens);
                    double gapWidth = gaps == 0 ? 0D : Math.Max(0D, lineWidth - line.Width) / gaps;
                    bool hasWord = false;
                    cursor = lineLeft;
                    for (int tokenIndex = 0; tokenIndex < tokens.Count; tokenIndex++) {
                        var token = tokens[tokenIndex];
                        double advance = token.Width;
                        if (token.IsWhitespace && hasWord && OfficeTextBlockRenderer.HasWordAfter(tokens, tokenIndex + 1)) advance += gapWidth;
                        Paint(token.Segment, token.Text, cursor, baseline, advance);
                        cursor += advance;
                        hasWord |= !token.IsWhitespace;
                    }
                } else {
                    foreach (OfficeRichTextSegment segment in line.Segments) {
                        Paint(segment, segment.Text, cursor, baseline, segment.Width);
                        cursor += segment.Width;
                    }
                }
                lineTop += lineHeight;
            }

            void Paint(OfficeRichTextSegment segment, string value, double x, double baseline, double advance) {
                double size = OfficeTextBlockRenderer.ResolveRichTextRenderedFontSize(segment);
                double renderedBaseline = OfficeTextBlockRenderer.ResolveRichTextRenderedBaseline(segment, baseline);
                var run = new PdfTextRun(value, segment.Bold, segment.Underline, ToPdfColor(segment.Color),
                    segment.Italic, segment.Strikethrough, size, ResolveDrawingTextFont(segment.FontFamily),
                    linkUri: segment.LinkUri, backgroundColor: ToPdfColor(segment.BackgroundColor),
                    fontFamily: segment.FontFamily, underlineStyle: segment.UnderlineStyle,
                    strikeStyle: segment.StrikethroughStyle, decorationColor: ToPdfColor(decorationColor));
                WriteDrawingPositionedRun(run, x, contentTopY - renderedBaseline, advance,
                    contentX, contentTopY - height, width, height);
            }
        }

        private void WriteDrawingPositionedRun(PdfTextRun run, double x, double baseline, double advance,
            double clipX, double clipY, double clipWidth, double clipHeight) {
            double size = run.FontSize ?? currentOpts.DefaultFontSize;
            var runs = new[] { run };
            var line = CreatePositionedTextLine(runs, size, size * 1.2D, currentOpts);
            double actualWidth = MeasureRichLineWidth(line.Lines[0], currentOpts);
            double scale = actualWidth > 0D && advance > 0D ? advance / actualWidth : 1D;
            void Paint() {
                WriteClippedRichParagraph(sb, new RichParagraphBlock(runs, PdfAlign.Left, null),
                    line.Lines, line.LineHeights, currentOpts, baseline, size, size * 1.2D,
                    currentPage!.Annotations, x + (clipX - x) / scale, clipY, clipWidth / scale,
                    clipHeight, x, Math.Max(0.001D, actualWidth), suppressActualText: _suppressCanvasActualTextChildren);
            }
            if (Math.Abs(scale - 1D) > 0.000001D) {
                RenderOpaqueEffectGroupInline(new OfficeTransform(scale, 0D, 0D, 1D, x * (1D - scale), 0D), Paint);
            } else {
                Paint();
            }
            MarkRichFonts(runs);
            pageDirty = true;
        }

        private PdfStandardFont ResolveDrawingTextFont(string? familyName) {
            if (!string.IsNullOrWhiteSpace(familyName) && PdfStandardFontMapper.TryMapFontFamily(familyName, out PdfStandardFont mapped)) {
                return ChooseNormal(mapped);
            }

            return string.IsNullOrWhiteSpace(familyName) ? ChooseNormal(currentOpts.DefaultFont) : PdfStandardFont.Helvetica;
        }

        private static PdfAlign MapDrawingTextAlignment(OfficeTextAlignment alignment) => alignment switch {
            OfficeTextAlignment.Center => PdfAlign.Center,
            OfficeTextAlignment.Right => PdfAlign.Right,
            OfficeTextAlignment.Justify => PdfAlign.Justify,
            _ => PdfAlign.Left
        };

        private static PdfTextBaseline MapDrawingTextBaseline(OfficeTextBaseline baseline) => baseline switch {
            OfficeTextBaseline.Superscript => PdfTextBaseline.Superscript,
            OfficeTextBaseline.Subscript => PdfTextBaseline.Subscript,
            _ => PdfTextBaseline.Normal
        };

        private static string StackTextElements(string value) {
            var builder = new StringBuilder(value.Length * 2);
            TextElementEnumerator enumerator = StringInfo.GetTextElementEnumerator(value);
            while (enumerator.MoveNext()) {
                string element = enumerator.GetTextElement();
                if (element == "\r") continue;
                if (builder.Length > 0 && builder[builder.Length - 1] != '\n' && element != "\n") builder.Append('\n');
                builder.Append(element);
            }

            return builder.ToString();
        }

    }
}
