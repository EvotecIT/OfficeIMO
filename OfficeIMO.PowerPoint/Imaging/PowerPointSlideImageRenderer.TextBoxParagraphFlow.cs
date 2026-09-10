using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointSlideImageRenderer {
        private static bool TryAddTextBoxParagraphFlow(
            OfficeDrawing drawing,
            PowerPointTextBox textBox,
            double left,
            double top,
            double width,
            double height,
            double textWidth,
            double textHeight,
            double marginLeft,
            double marginTop,
            double rotation,
            double rotationCenterX,
            double rotationCenterY,
            bool flipHorizontal,
            bool flipVertical,
            PowerPointShapeBoundsMapping mapping,
            A.ColorScheme? colorScheme,
            List<OfficeImageExportDiagnostic> diagnostics, Func<string?, double, string?, OfficeFontStyle, double> measure, string? defaultFontFamily) {
            List<PowerPointParagraph> paragraphs = GetVisibleTextBoxParagraphs(textBox);
            if (!ShouldRenderTextBoxParagraphFlow(paragraphs)) {
                return false;
            }

            var numberingState = new Dictionary<int, int>();
            List<PowerPointParagraphDrawing> paragraphDrawings = CreateTextBoxParagraphDrawings(textBox, paragraphs, numberingState, textWidth, mapping, colorScheme, measure, defaultFontFamily);
            double flowHeight = paragraphDrawings.Sum(paragraph => paragraph.TotalHeight);
            double currentY = top + marginTop + ResolveTextBoxVerticalOffset(textBox.TextVerticalAlignment, textHeight, flowHeight);
            double contentBottom = top + marginTop + textHeight;

            for (int i = 0; i < paragraphDrawings.Count; i++) {
                PowerPointParagraphDrawing paragraph = paragraphDrawings[i];
                currentY += paragraph.SpaceBefore;
                double visibleHeight = Math.Min(paragraph.Height, Math.Max(0D, contentBottom - currentY));
                bool clipped = visibleHeight < paragraph.Height;
                if (clipped) {
                    diagnostics.Add(new OfficeImageExportDiagnostic(OfficeImageExportDiagnosticSeverity.Warning,
                        "POWERPOINT_TEXT_OVERFLOW", "Clipped PowerPoint paragraph content at the text frame boundary.", DescribeShape(textBox), OfficeConversionLossKind.Omission));
                }
                if (visibleHeight <= 0D) return true;

                if (paragraph.RichRuns.Count > 0) {
                    drawing.AddRichText(
                        paragraph.RichRuns,
                        left + marginLeft,
                        currentY,
                        textWidth,
                        visibleHeight,
                        paragraph.Alignment,
                        paragraph.LineHeight,
                        rotationDegrees: rotation,
                        rotationCenterX: rotationCenterX,
                        rotationCenterY: rotationCenterY,
                        wrapText: true,
                        flipHorizontal: flipHorizontal,
                        flipVertical: flipVertical,
                        paragraphIndent: paragraph.Indent);
                } else {
                    drawing.AddText(
                        paragraph.Text,
                        left + marginLeft,
                        currentY,
                        textWidth,
                        visibleHeight,
                        paragraph.Font,
                        paragraph.Color,
                        paragraph.Alignment,
                        paragraph.LineHeight,
                        rotationDegrees: rotation,
                        rotationCenterX: rotationCenterX,
                        rotationCenterY: rotationCenterY,
                        wrapText: true,
                        flipHorizontal: flipHorizontal,
                        flipVertical: flipVertical,
                        paragraphIndent: paragraph.Indent);
                }

                if (clipped) return true;
                currentY += paragraph.Height + paragraph.SpaceAfter;
            }

            return true;
        }

        private static List<PowerPointParagraph> GetVisibleTextBoxParagraphs(PowerPointTextBox textBox) =>
            textBox.Paragraphs
                .Where(paragraph => paragraph.InlineNodes.Any(node => !string.IsNullOrEmpty(node.Text)) || !string.IsNullOrEmpty(paragraph.BulletCharacter) || paragraph.IsNumbered)
                .ToList();

        private static bool ShouldRenderTextBoxParagraphFlow(IReadOnlyList<PowerPointParagraph> paragraphs) =>
            paragraphs.Any(paragraph => !string.IsNullOrEmpty(paragraph.BulletCharacter) || paragraph.IsNumbered) ||
            (paragraphs.Count > 1 &&
                paragraphs.Any(paragraph =>
                    paragraph.Alignment != paragraphs[0].Alignment || paragraph.SpaceBeforePoints.HasValue ||
                    paragraph.SpaceAfterPoints.HasValue ||
                    paragraph.LineSpacingPoints.HasValue ||
                    paragraph.LineSpacingMultiplier.HasValue));

        private static List<PowerPointParagraphDrawing> CreateTextBoxParagraphDrawings(
            PowerPointTextBox textBox,
            IReadOnlyList<PowerPointParagraph> paragraphs,
            Dictionary<int, int> numberingState,
            double textWidth,
            PowerPointShapeBoundsMapping mapping,
            A.ColorScheme? colorScheme, Func<string?, double, string?, OfficeFontStyle, double> measure, string? defaultFontFamily) {
            var results = new List<PowerPointParagraphDrawing>(paragraphs.Count);
            for (int i = 0; i < paragraphs.Count; i++) {
                PowerPointParagraph paragraph = paragraphs[i];
                string? marker = CreateParagraphMarker(paragraph, numberingState);
                OfficeTextParagraphIndent indent = CreateParagraphIndent(paragraph, mapping);
                OfficeTextAlignment alignment = MapTextAlignment(paragraph.Alignment);
                List<OfficeRichTextRun> richRuns = CreateParagraphRichTextRuns(textBox, paragraph, marker, colorScheme, mapping, defaultFontFamily);
                double maxFontSize = richRuns.Count == 0
                    ? ResolveParagraphFont(textBox, paragraph, mapping, defaultFontFamily).Size
                    : richRuns.Max(run => run.FontSize);
                double lineHeight = ResolveParagraphLineHeight(paragraph, maxFontSize, mapping);
                double height;
                if (ShouldRenderParagraphRichText(richRuns, marker)) {
                    height = EstimateParagraphRichTextHeight(richRuns, maxFontSize, lineHeight, textWidth, indent, measure);
                    results.Add(PowerPointParagraphDrawing.FromRichText(paragraph, richRuns, alignment, indent, lineHeight, height, mapping));
                } else {
                    string text = CreateParagraphPlainText(paragraph, marker);
                    OfficeFontInfo font = richRuns.Count == 1 ? new OfficeFontInfo(richRuns[0].FontFamily, richRuns[0].FontSize, richRuns[0].FontStyle) : ResolveParagraphFont(textBox, paragraph, mapping, defaultFontFamily);
                    height = EstimateParagraphTextHeight(text, font, lineHeight, textWidth, indent, measure);
                    results.Add(PowerPointParagraphDrawing.FromText(paragraph, text, font, ResolveParagraphTextColor(textBox, paragraph, colorScheme), alignment, indent, lineHeight, height, mapping));
                }
            }

            return results;
        }

        private static bool ShouldRenderParagraphRichText(IReadOnlyList<OfficeRichTextRun> runs, string? marker) =>
            !string.IsNullOrEmpty(marker) || ShouldRenderRichText(runs);

        private static List<OfficeRichTextRun> CreateParagraphRichTextRuns(
            PowerPointTextBox textBox,
            PowerPointParagraph paragraph,
            string? marker,
            A.ColorScheme? colorScheme,
            PowerPointShapeBoundsMapping mapping, string? defaultFontFamily) {
            IReadOnlyList<PowerPointParagraphInline> inlineNodes = paragraph.InlineNodes;
            PowerPointTextRun? firstRun = inlineNodes.FirstOrDefault(node => node.Run != null)?.Run;
            var richRuns = new List<OfficeRichTextRun>();
            if (!string.IsNullOrEmpty(marker)) {
                richRuns.Add(CreateRichTextRun(marker!, firstRun, textBox, paragraph, colorScheme, mapping, markerRun: true, defaultFontFamily: defaultFontFamily));
            }

            for (int i = 0; i < inlineNodes.Count; i++) {
                PowerPointParagraphInline inline = inlineNodes[i];
                if (!string.IsNullOrEmpty(inline.Text)) {
                    richRuns.AddRange(CreateEffectiveTextRuns(inline.Text, inline.Run, textBox, paragraph, colorScheme, mapping, defaultFontFamily));
                }
            }

            return richRuns;
        }

        private static string CreateParagraphPlainText(PowerPointParagraph paragraph, string? marker) {
            string text = string.Concat(paragraph.InlineNodes.Select(node =>
                node.Run == null ? node.Text : ResolvePowerPointDisplayText(node.Text, node.Run, paragraph)));
            return string.IsNullOrEmpty(marker) ? text : marker + text;
        }

        private static OfficeFontInfo ResolveParagraphFont(PowerPointTextBox textBox, PowerPointParagraph paragraph, PowerPointShapeBoundsMapping mapping, string? defaultFontFamily) {
            PowerPointTextRun? firstRun = paragraph.InlineNodes.FirstOrDefault(node => node.Run != null && !string.IsNullOrEmpty(node.Text))?.Run
                ?? paragraph.InlineNodes.FirstOrDefault(node => node.Run != null)?.Run;
            OfficeFontStyle style = OfficeFontStyle.Regular;
            if (firstRun?.Bold == true || textBox.Bold) {
                style |= OfficeFontStyle.Bold;
            }

            if (firstRun?.Italic == true || textBox.Italic) {
                style |= OfficeFontStyle.Italic;
            }

            if (firstRun?.Underline == true) {
                style |= OfficeFontStyle.Underline;
            }

            if (firstRun?.Strikethrough == true) {
                style |= OfficeFontStyle.Strikethrough;
            }

            return new OfficeFontInfo(firstRun?.FontName ?? textBox.FontName ?? ResolveTextBoxFallbackFont(textBox, defaultFontFamily), mapping.MapFontSize(firstRun?.FontSize ?? textBox.FontSize ?? 18), style);
        }

        private static OfficeColor ResolveParagraphTextColor(PowerPointTextBox textBox, PowerPointParagraph paragraph, A.ColorScheme? colorScheme) {
            PowerPointTextRun? firstRun = paragraph.InlineNodes.FirstOrDefault(node => node.Run != null && !string.IsNullOrEmpty(node.Text))?.Run
                ?? paragraph.InlineNodes.FirstOrDefault(node => node.Run != null)?.Run;
            return ResolveTextRunColor(firstRun, textBox, colorScheme);
        }

        private static double ResolveParagraphLineHeight(PowerPointParagraph paragraph, double fontSize, PowerPointShapeBoundsMapping mapping) {
            if (paragraph.LineSpacingPoints.HasValue) {
                return Math.Max(1D, mapping.MapVerticalLength(paragraph.LineSpacingPoints.Value));
            }

            if (paragraph.LineSpacingMultiplier.HasValue) {
                return Math.Max(1D, fontSize * paragraph.LineSpacingMultiplier.Value);
            }

            return Math.Max(1D, fontSize * 1.2D);
        }

        private static double EstimateParagraphTextHeight(string text, OfficeFontInfo font, double lineHeight, double textWidth, OfficeTextParagraphIndent indent, Func<string?, double, string?, OfficeFontStyle, double>? styledMeasure = null) {
            styledMeasure ??= OfficeDrawingTextLayout.CreateMetrics(null).MeasureText;
            Func<string?, double, double> measure = (value, size) => styledMeasure(value, size, font.FamilyName, font.Style);
            OfficeTextBlockLayout layout = OfficeTextLayoutEngine.LayoutTextBlock(
                text,
                font.Size,
                textWidth,
                double.MaxValue,
                Math.Max(1D, lineHeight / Math.Max(1D, font.Size)),
                Math.Min(6D, font.Size),
                measure,
                wrap: true,
                paragraphIndent: indent);
            return Math.Max(lineHeight, layout.Height);
        }

        private static double EstimateParagraphRichTextHeight(IReadOnlyList<OfficeRichTextRun> runs, double maxFontSize, double lineHeight, double textWidth, OfficeTextParagraphIndent indent,
            Func<string?, double, string?, OfficeFontStyle, double> measure) {
            OfficeRichTextBlockLayout layout = OfficeTextLayoutEngine.LayoutStyledRichTextBlock(
                runs, textWidth, double.MaxValue, Math.Max(1D, lineHeight / Math.Max(1D, maxFontSize)),
                measure, wrap: true, minimumFontSize: Math.Min(6D, maxFontSize), paragraphIndent: indent);
            return Math.Max(lineHeight, layout.Height);
        }

        private static double ResolveTextBoxVerticalOffset(PowerPointTextVerticalAlignment? alignment, double textHeight, double flowHeight) {
            double extraHeight = Math.Max(0D, textHeight - flowHeight);
            if (alignment == PowerPointTextVerticalAlignment.Center) {
                return extraHeight / 2D;
            }

            if (alignment == PowerPointTextVerticalAlignment.Bottom) {
                return extraHeight;
            }

            return 0D;
        }

        private readonly struct PowerPointParagraphDrawing {
            internal PowerPointParagraphDrawing(
                string text,
                IReadOnlyList<OfficeRichTextRun> richRuns,
                OfficeFontInfo font,
                OfficeColor color,
                OfficeTextAlignment alignment,
                OfficeTextParagraphIndent indent,
                double lineHeight,
                double height,
                double spaceBefore,
                double spaceAfter) {
                Text = text;
                RichRuns = richRuns;
                Font = font;
                Color = color;
                Alignment = alignment;
                Indent = indent;
                LineHeight = lineHeight;
                Height = height;
                SpaceBefore = spaceBefore;
                SpaceAfter = spaceAfter;
            }

            internal string Text { get; }

            internal IReadOnlyList<OfficeRichTextRun> RichRuns { get; }

            internal OfficeFontInfo Font { get; }

            internal OfficeColor Color { get; }

            internal OfficeTextAlignment Alignment { get; }

            internal OfficeTextParagraphIndent Indent { get; }

            internal double LineHeight { get; }

            internal double Height { get; }

            internal double SpaceBefore { get; }

            internal double SpaceAfter { get; }

            internal double TotalHeight => SpaceBefore + Height + SpaceAfter;

            internal static PowerPointParagraphDrawing FromText(
                PowerPointParagraph paragraph,
                string text,
                OfficeFontInfo font,
                OfficeColor color,
                OfficeTextAlignment alignment,
                OfficeTextParagraphIndent indent,
                double lineHeight,
                double height,
                PowerPointShapeBoundsMapping mapping) =>
                new PowerPointParagraphDrawing(
                    text,
                    Array.Empty<OfficeRichTextRun>(),
                    font,
                    color,
                    alignment,
                    indent,
                    lineHeight,
                    height,
                    Math.Max(0D, mapping.MapVerticalLength(paragraph.SpaceBeforePoints ?? 0D)),
                    Math.Max(0D, mapping.MapVerticalLength(paragraph.SpaceAfterPoints ?? 0D)));

            internal static PowerPointParagraphDrawing FromRichText(
                PowerPointParagraph paragraph,
                IReadOnlyList<OfficeRichTextRun> richRuns,
                OfficeTextAlignment alignment,
                OfficeTextParagraphIndent indent,
                double lineHeight,
                double height,
                PowerPointShapeBoundsMapping mapping) =>
                new PowerPointParagraphDrawing(
                    string.Empty,
                    richRuns,
                    OfficeFontInfo.Default,
                    OfficeColor.Black,
                    alignment,
                    indent,
                    lineHeight,
                    height,
                    Math.Max(0D, mapping.MapVerticalLength(paragraph.SpaceBeforePoints ?? 0D)),
                    Math.Max(0D, mapping.MapVerticalLength(paragraph.SpaceAfterPoints ?? 0D)));
        }
    }
}
