using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointSlideImageRenderer {
        /// <summary>
        /// Applies DrawingML fitting before placement. Fixed frames retain overflowing
        /// paragraphs; shape fitting grows only the exported frame, without editing the deck.
        /// </summary>
        private static OfficeDrawing CreateFittedTextBoxDrawing(PowerPointTextBox textBox, double width, double height,
            List<OfficeImageExportDiagnostic> diagnostics, PowerPointShapeBoundsMapping mapping, A.ColorScheme? colors,
            Func<string?, double, string?, OfficeFontStyle, double> measure, string? defaultFontFamily,
            CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            if (string.IsNullOrEmpty(ResolvePowerPointDisplayText(textBox))) return new OfficeDrawing(width, height);

            ReportCharacterSpacing(textBox, diagnostics);
            AddSmallCapsApproximationDiagnosticIfNeeded(textBox, diagnostics);
            AddWavyDoubleUnderlineApproximationDiagnosticIfNeeded(textBox, diagnostics);
            AddBaselineApproximationDiagnosticIfNeeded(textBox, diagnostics);

            double marginLeft = mapping.MapHorizontalLength(textBox.TextMarginLeftPoints ?? PowerPointTextDefaults.DefaultHorizontalInsetPoints);
            double marginRight = mapping.MapHorizontalLength(textBox.TextMarginRightPoints ?? PowerPointTextDefaults.DefaultHorizontalInsetPoints);
            double marginTop = mapping.MapVerticalLength(textBox.TextMarginTopPoints ?? PowerPointTextDefaults.DefaultVerticalInsetPoints);
            double marginBottom = mapping.MapVerticalLength(textBox.TextMarginBottomPoints ?? PowerPointTextDefaults.DefaultVerticalInsetPoints);
            if (marginLeft + marginRight >= width || marginTop + marginBottom >= height) {
                double horizontalScale = Math.Min(1D, width * 0.99D / Math.Max(0.01D, marginLeft + marginRight));
                double verticalScale = Math.Min(1D, height * 0.99D / Math.Max(0.01D, marginTop + marginBottom));
                marginLeft *= horizontalScale; marginRight *= horizontalScale;
                marginTop *= verticalScale; marginBottom *= verticalScale;
                diagnostics.Add(new OfficeImageExportDiagnostic(OfficeImageExportDiagnosticSeverity.Warning,
                    "powerpoint-text-padding", "Reduced PowerPoint text box margins because they leave no renderable text area.",
                    DescribeShape(textBox), OfficeConversionLossKind.Approximation));
            }
            double textWidth = width - marginLeft - marginRight;
            double textHeight = height - marginTop - marginBottom;
            bool wrap = textBox.TextBody?.GetFirstChild<A.BodyProperties>()?.Wrap?.Value != A.TextWrappingValues.None;
            List<PowerPointParagraph> paragraphs = textBox.Paragraphs.ToList();
            PowerPointTextAutoFitOptions? options = textBox.TextAutoFit == PowerPointTextAutoFit.Normal
                ? textBox.TextAutoFitOptions : null;
            double fontScale = options?.FontScalePercent / 100D ?? 1D;
            double reduction = options?.LineSpaceReductionPercent / 100D ?? 0D;
            // Drawing fonts must remain positive, including an authored zero scale.
            if (fontScale <= 0D) {
                fontScale = 0.001D;
                diagnostics.Add(new OfficeImageExportDiagnostic(OfficeImageExportDiagnosticSeverity.Warning,
                    "powerpoint-text-scale-approximated", "Approximated a zero text font scale with a positive rendering size.",
                    DescribeShape(textBox), OfficeConversionLossKind.Approximation));
            }

            double MeasureText(string? text, double size, string? family, OfficeFontStyle style) {
                cancellationToken.ThrowIfCancellationRequested();
                return measure(text, size, family, style);
            }
            List<PowerPointParagraphDrawing> Measure(double scale) => CreateTextBoxParagraphDrawings(textBox, paragraphs,
                new PowerPointParagraphNumberingState(), textWidth, mapping, colors, MeasureText, defaultFontFamily,
                scale, reduction, wrap, cancellationToken);
            List<PowerPointParagraphDrawing> measured = Measure(fontScale);
            if (textBox.TextAutoFit == PowerPointTextAutoFit.Normal && options?.FontScalePercent == null) {
                double maximumFontSize = measured.Max(paragraph => paragraph.RichRuns.Count == 0
                    ? paragraph.Font.Size : paragraph.RichRuns.Max(run => run.FontSize));
                fontScale = OfficeTextLayoutEngine.ResolveFrameFitScale(Math.Min(1D, 1D / maximumFontSize),
                    scale => TextBoxFlowFits(Measure(scale), textWidth, textHeight), cancellationToken);
                measured = Measure(fontScale);
            }

            double flowHeight = measured.Sum(paragraph => paragraph.TotalHeight);
            double paintedBottom = TextBoxPaintedBottom(measured);
            bool growsShape = textBox.TextAutoFit == PowerPointTextAutoFit.Shape;
            if (growsShape) {
                height = Math.Max(height, marginTop + Math.Max(flowHeight, paintedBottom) + marginBottom);
                textHeight = height - marginTop - marginBottom;
                if (!wrap) {
                    width = Math.Max(width, marginLeft + measured.Max(paragraph => paragraph.MeasuredWidth) + marginRight);
                    textWidth = width - marginLeft - marginRight;
                }
            }
            bool exceedsFrame = paintedBottom > textHeight + 0.001D ||
                measured.Any(paragraph => paragraph.HasContent && paragraph.MeasuredWidth > textWidth + 0.001D);
            if (exceedsFrame) {
                diagnostics.Add(new OfficeImageExportDiagnostic(OfficeImageExportDiagnosticSeverity.Warning,
                    "powerpoint-text-overflow", "PowerPoint text exceeds the text frame and is rendered beyond its boundary without truncation.",
                    DescribeShape(textBox), OfficeConversionLossKind.None));
            }
            if (measured.Any(paragraph => paragraph.Clipped)) {
                diagnostics.Add(new OfficeImageExportDiagnostic(OfficeImageExportDiagnosticSeverity.Warning,
                    "powerpoint-text-layout-limit", "PowerPoint text exceeded the shared text layout safety limits; some content could not be laid out.",
                    DescribeShape(textBox), OfficeConversionLossKind.Omission));
            }

            var drawing = new OfficeDrawing(width, height);
            double currentY = marginTop + ResolveTextBoxVerticalOffset(textBox.TextVerticalAlignment, textHeight, flowHeight);
            double centerX = width / 2D, centerY = height / 2D;
            foreach (PowerPointParagraphDrawing paragraph in measured) {
                cancellationToken.ThrowIfCancellationRequested();
                currentY += paragraph.SpaceBefore;
                if (paragraph.HasContent) {
                    // Leading advances the next origin. It must not clip the current glyphs.
                    double renderHeight = Math.Max(paragraph.Height, paragraph.PaintedHeight) + 0.01D;
                    double renderWidth = (wrap ? Math.Max(textWidth, paragraph.MeasuredWidth) : textWidth) + 0.000001D;
                    var content = new OfficeDrawing(renderWidth, renderHeight);
                    if (!wrap || paragraph.AbsoluteLineHeight) {
                        // Keep fixed point leading exact across every visual line.
                        // Unwrapped text may overhang its authored alignment anchor.
                        content.AddRichTextParagraphs(new[] { new OfficeRichTextParagraph(paragraph.RichRuns,
                            paragraph.Alignment, lineHeight: paragraph.LineHeight, indent: paragraph.Indent) },
                            0D, 0D, renderWidth, renderHeight, wrapText: wrap);
                    } else if (paragraph.RichRuns.Count > 0) {
                        content.AddRichText(paragraph.RichRuns, 0D, 0D, renderWidth, renderHeight, paragraph.Alignment,
                            paragraph.LineHeight, paragraphIndent: paragraph.Indent);
                    } else {
                        content.AddText(paragraph.Text, 0D, 0D, renderWidth, renderHeight, paragraph.Font,
                            paragraph.Color, paragraph.Alignment, paragraph.LineHeight, wrapText: true, paragraphIndent: paragraph.Indent);
                    }
                    // Flatten rather than wrap in a shape-sized clip; overflow stays visible on the slide.
                    drawing.AddDrawingForClippedRendering(content, marginLeft, currentY,
                        new OfficeImageFrameTransform(textBox.Rotation ?? 0D, centerX, centerY,
                            textBox.HorizontalFlip == true, textBox.VerticalFlip == true));
                }
                currentY += paragraph.Height + paragraph.SpaceAfter;
            }
            return drawing;
        }

        private static bool TextBoxFlowFits(IReadOnlyList<PowerPointParagraphDrawing> paragraphs, double width, double height) =>
            paragraphs.Sum(paragraph => paragraph.TotalHeight) <= height + 0.001D &&
            TextBoxPaintedBottom(paragraphs) <= height + 0.001D &&
            paragraphs.All(paragraph => !paragraph.Clipped && paragraph.MeasuredWidth <= width + 0.001D);

        private static double TextBoxPaintedBottom(IReadOnlyList<PowerPointParagraphDrawing> paragraphs) {
            double currentY = 0D, bottom = 0D;
            foreach (PowerPointParagraphDrawing paragraph in paragraphs) {
                currentY += paragraph.SpaceBefore;
                if (paragraph.HasContent) bottom = Math.Max(bottom, currentY + Math.Max(paragraph.Height, paragraph.PaintedHeight));
                currentY += paragraph.Height + paragraph.SpaceAfter;
            }
            return bottom;
        }

        private static List<OfficeRichTextRun> ScaleTextBoxRuns(IReadOnlyList<OfficeRichTextRun> runs, double scale,
            CancellationToken cancellationToken = default) {
            var scaled = new List<OfficeRichTextRun>(runs.Count);
            foreach (OfficeRichTextRun run in runs) {
                cancellationToken.ThrowIfCancellationRequested();
                scaled.Add(new OfficeRichTextRun(run.Text, run.FontSize * scale, run.Color, run.Bold, run.Italic,
                    run.Underline, run.FontFamily, run.Strikethrough, run.BackgroundColor, run.UnderlineStyle,
                    run.StrikethroughStyle, run.Baseline, run.ParagraphIndent) {
                    LinkUri = run.LinkUri, LinkDestinationName = run.LinkDestinationName
                });
            }
            return scaled;
        }
    }
}
