using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void DrawDrawingTextAt(OfficeDrawingText text, double originX, double originTopY) {
            if (string.IsNullOrEmpty(text.Text)) return;

            string value = text.StackedText ? StackTextElements(text.Text) : text.Text;
            var run = new OfficeRichTextRun(
                value,
                text.Font.Size,
                text.Color ?? OfficeColor.Black,
                text.Font.IsBold,
                text.Font.IsItalic,
                text.Font.IsUnderline,
                text.Font.FamilyName,
                text.Font.IsStrikethrough);
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
            DrawDrawingRichTextAt(richText, originX, originTopY);
        }

        private void DrawDrawingRichTextAt(OfficeDrawingRichText text, double originX, double originTopY) {
            if (text.Runs.Count == 0 || string.IsNullOrEmpty(text.PlainText)) return;

            void DrawContent() => DrawDrawingRichTextCore(text, originX + text.X, originTopY - text.Y);
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

        private void DrawDrawingRichTextCore(OfficeDrawingRichText text, double frameX, double frameTopY) {
            double contentX = frameX + text.Padding.Left;
            double contentTopY = frameTopY - text.Padding.Top;
            double contentWidth = text.Width - text.Padding.Horizontal;
            double contentHeight = text.Height - text.Padding.Vertical;
            if (contentWidth <= 0D || contentHeight <= 0D) return;

            DrawingRichTextLayout layout = CreateDrawingRichTextLayout(text, contentWidth, contentHeight);
            if (layout.Lines.Count == 0) return;

            double contentUsedHeight = MeasureRichLinesHeight(layout.LineHeights, layout.Lines.Count, layout.Leading);
            double verticalOffset = text.VerticalAlignment switch {
                OfficeTextVerticalAlignment.Center => Math.Max(0D, (contentHeight - contentUsedHeight) / 2D),
                OfficeTextVerticalAlignment.Bottom => Math.Max(0D, contentHeight - contentUsedHeight),
                _ => 0D
            };
            var block = new RichParagraphBlock(layout.Runs, MapDrawingTextAlignment(text.Alignment), null);
            PdfStandardFont baseFont = ChooseNormal(currentOpts.DefaultFont);
            WriteClippedRichParagraph(
                sb,
                block,
                layout.Lines,
                layout.LineHeights,
                currentOpts,
                FirstTextBaselineFromTop(baseFont, layout.BaseFontSize, contentTopY - verticalOffset),
                layout.BaseFontSize,
                layout.Leading,
                currentPage!.Annotations,
                contentX,
                contentTopY - contentHeight,
                contentWidth,
                contentHeight,
                contentX,
                contentWidth,
                structureType: null,
                markedContentId: null,
                structurePage: null,
                lineXOffsets: layout.LineOffsets,
                lineWidths: layout.LineWidths);
            MarkRichFonts(layout.Runs);
            pageDirty = true;
        }

        private DrawingRichTextLayout CreateDrawingRichTextLayout(OfficeDrawingRichText text, double width, double height) {
            double maximumFontSize = text.Runs.Max(run => run.FontSize);
            DrawingRichTextLayout fullSize = BuildDrawingRichTextLayout(text, width, 1D);
            if (!text.ShrinkToFit || DrawingRichTextFits(fullSize, height)) return fullSize;

            double minimumScale = maximumFontSize > 6D ? 6D / maximumFontSize : 1D;
            DrawingRichTextLayout minimum = BuildDrawingRichTextLayout(text, width, minimumScale);
            if (!DrawingRichTextFits(minimum, height)) return minimum;

            double low = minimumScale;
            double high = 1D;
            DrawingRichTextLayout best = minimum;
            for (int iteration = 0; iteration < 24; iteration++) {
                double candidateScale = (low + high) / 2D;
                DrawingRichTextLayout candidate = BuildDrawingRichTextLayout(text, width, candidateScale);
                if (DrawingRichTextFits(candidate, height)) {
                    best = candidate;
                    low = candidateScale;
                } else {
                    high = candidateScale;
                }
            }

            return best;
        }

        private DrawingRichTextLayout BuildDrawingRichTextLayout(OfficeDrawingRichText text, double width, double scale) {
            List<PdfTextRun> runs = text.Runs.Select(run => new PdfTextRun(
                run.Text,
                run.Bold,
                run.Underline,
                ToPdfColor(run.Color),
                run.Italic,
                run.Strikethrough,
                run.FontSize * scale,
                ResolveDrawingTextFont(run.FontFamily),
                backgroundColor: ToPdfColor(run.BackgroundColor),
                fontFamily: run.FontFamily,
                baseline: MapDrawingTextBaseline(run.Baseline),
                underlineStyle: run.UnderlineStyle,
                strikeStyle: run.StrikethroughStyle)).ToList();

            double baseFontSize = Math.Max(0.001D, text.Runs.Max(run => run.FontSize) * scale);
            double leading = text.LineHeight.HasValue ? text.LineHeight.Value * scale : baseFontSize * 1.2D;
            double firstOffset = Math.Min(width, text.ParagraphIndent.FirstLineOffset);
            double continuationOffset = Math.Min(width, text.ParagraphIndent.ContinuationLineOffset);
            double continuationWidth = Math.Max(0.001D, width - continuationOffset);
            double firstWidth = Math.Max(0.001D, width - firstOffset);
            double wrapWidth = text.WrapText ? continuationWidth : 1_000_000_000D;
            double? firstLineWidth = text.WrapText ? firstWidth : null;
            double? firstLineOrigin = text.WrapText ? firstOffset - continuationOffset : null;
            var wrap = WrapRichRunsCoreWithFirstLineOrigin(
                runs,
                wrapWidth,
                baseFontSize,
                ChooseNormal(currentOpts.DefaultFont),
                leading,
                firstLineWidth,
                firstLineOrigin,
                DefaultParagraphTabStopWidth,
                currentOpts);
            var offsets = new List<double>(wrap.Lines.Count);
            var widths = new List<double>(wrap.Lines.Count);
            for (int index = 0; index < wrap.Lines.Count; index++) {
                offsets.Add(index == 0 ? firstOffset : continuationOffset);
                widths.Add(index == 0 ? firstWidth : continuationWidth);
            }

            return new DrawingRichTextLayout(runs, wrap.Lines, wrap.LineHeights, offsets, widths, baseFontSize, leading);
        }

        private bool DrawingRichTextFits(DrawingRichTextLayout layout, double height) {
            if (MeasureRichLinesHeight(layout.LineHeights, layout.Lines.Count, layout.Leading) > height + 0.001D) return false;
            for (int lineIndex = 0; lineIndex < layout.Lines.Count; lineIndex++) {
                double lineWidth = 0D;
                foreach (RichSeg segment in layout.Lines[lineIndex]) {
                    if (segment.LeadingSpace) {
                        lineWidth += segment.LeadingAdvance > 0D
                            ? segment.LeadingAdvance
                            : MeasureRichText(" ", segment.Font, segment.NamedFont, segment.FontSize, segment.Baseline, currentOpts);
                    }

                    lineWidth += MeasureRichSegment(segment, currentOpts);
                }

                if (lineWidth > layout.LineWidths[lineIndex] + 0.001D) return false;
            }

            return true;
        }

        private PdfStandardFont ResolveDrawingTextFont(string? familyName) {
            if (!string.IsNullOrWhiteSpace(familyName) && PdfStandardFontMapper.TryMapFontFamily(familyName, out PdfStandardFont mapped)) {
                return ChooseNormal(mapped);
            }

            return ChooseNormal(currentOpts.DefaultFont);
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

        private sealed class DrawingRichTextLayout {
            internal DrawingRichTextLayout(
                IReadOnlyList<PdfTextRun> runs,
                List<List<RichSeg>> lines,
                List<double> lineHeights,
                IReadOnlyList<double> lineOffsets,
                IReadOnlyList<double> lineWidths,
                double baseFontSize,
                double leading) {
                Runs = runs;
                Lines = lines;
                LineHeights = lineHeights;
                LineOffsets = lineOffsets;
                LineWidths = lineWidths;
                BaseFontSize = baseFontSize;
                Leading = leading;
            }

            internal IReadOnlyList<PdfTextRun> Runs { get; }
            internal List<List<RichSeg>> Lines { get; }
            internal List<double> LineHeights { get; }
            internal IReadOnlyList<double> LineOffsets { get; }
            internal IReadOnlyList<double> LineWidths { get; }
            internal double BaseFontSize { get; }
            internal double Leading { get; }
        }
    }
}
