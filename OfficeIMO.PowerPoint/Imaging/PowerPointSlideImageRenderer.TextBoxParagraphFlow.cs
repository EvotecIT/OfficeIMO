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
            List<OfficeImageExportDiagnostic> diagnostics) {
            List<PowerPointParagraph> paragraphs = GetVisibleTextBoxParagraphs(textBox);
            if (!ShouldRenderTextBoxParagraphFlow(paragraphs)) {
                return false;
            }

            var numberingState = new Dictionary<int, int>();
            List<PowerPointParagraphDrawing> paragraphDrawings = CreateTextBoxParagraphDrawings(textBox, paragraphs, numberingState, textWidth, mapping, colorScheme);
            double flowHeight = paragraphDrawings.Sum(paragraph => paragraph.TotalHeight);
            double currentY = top + marginTop + ResolveTextBoxVerticalOffset(textBox.TextVerticalAlignment, textHeight, flowHeight);
            double contentBottom = top + marginTop + textHeight;

            for (int i = 0; i < paragraphDrawings.Count; i++) {
                PowerPointParagraphDrawing paragraph = paragraphDrawings[i];
                currentY += paragraph.SpaceBefore;
                if (currentY + paragraph.Height > contentBottom) {
                    AddUnsupportedShapeDiagnostic(diagnostics, textBox, "Skipped PowerPoint text box paragraph content because it does not fit within the text frame.");
                    return true;
                }

                if (paragraph.RichRuns.Count > 0) {
                    drawing.AddRichText(
                        paragraph.RichRuns,
                        left + marginLeft,
                        currentY,
                        textWidth,
                        paragraph.Height,
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
                        paragraph.Height,
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

                currentY += paragraph.Height + paragraph.SpaceAfter;
            }

            return true;
        }

        private static List<PowerPointParagraph> GetVisibleTextBoxParagraphs(PowerPointTextBox textBox) =>
            textBox.Paragraphs
                .Where(paragraph => paragraph.Runs.Any(run => !string.IsNullOrEmpty(run.Text)) || !string.IsNullOrEmpty(paragraph.BulletCharacter) || paragraph.IsNumbered)
                .ToList();

        private static bool ShouldRenderTextBoxParagraphFlow(IReadOnlyList<PowerPointParagraph> paragraphs) =>
            paragraphs.Any(paragraph => !string.IsNullOrEmpty(paragraph.BulletCharacter) || paragraph.IsNumbered) ||
            (paragraphs.Count > 1 &&
                paragraphs.Any(paragraph =>
                    paragraph.SpaceBeforePoints.HasValue ||
                    paragraph.SpaceAfterPoints.HasValue ||
                    paragraph.LineSpacingPoints.HasValue ||
                    paragraph.LineSpacingMultiplier.HasValue));

        private static List<PowerPointParagraphDrawing> CreateTextBoxParagraphDrawings(
            PowerPointTextBox textBox,
            IReadOnlyList<PowerPointParagraph> paragraphs,
            Dictionary<int, int> numberingState,
            double textWidth,
            PowerPointShapeBoundsMapping mapping,
            A.ColorScheme? colorScheme) {
            var results = new List<PowerPointParagraphDrawing>(paragraphs.Count);
            for (int i = 0; i < paragraphs.Count; i++) {
                PowerPointParagraph paragraph = paragraphs[i];
                string? marker = CreateParagraphMarker(paragraph, numberingState);
                OfficeTextParagraphIndent indent = CreateParagraphIndent(paragraph, mapping);
                OfficeTextAlignment alignment = MapTextAlignment(paragraph.Alignment);
                List<OfficeRichTextRun> richRuns = CreateParagraphRichTextRuns(textBox, paragraph, marker, colorScheme, mapping);
                double maxFontSize = richRuns.Count == 0
                    ? ResolveParagraphFont(textBox, paragraph, mapping).Size
                    : richRuns.Max(run => run.FontSize);
                double lineHeight = ResolveParagraphLineHeight(paragraph, maxFontSize, mapping);
                double height;
                if (ShouldRenderParagraphRichText(richRuns, marker)) {
                    height = EstimateParagraphRichTextHeight(richRuns, maxFontSize, lineHeight, textWidth, indent);
                    results.Add(PowerPointParagraphDrawing.FromRichText(paragraph, richRuns, alignment, indent, lineHeight, height, mapping));
                } else {
                    string text = CreateParagraphPlainText(paragraph, marker);
                    OfficeFontInfo font = ResolveParagraphFont(textBox, paragraph, mapping);
                    height = EstimateParagraphTextHeight(text, font, lineHeight, textWidth, indent);
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
            PowerPointShapeBoundsMapping mapping) {
            IReadOnlyList<PowerPointTextRun> runs = paragraph.Runs;
            PowerPointTextRun? firstRun = runs.Count > 0 ? runs[0] : null;
            var richRuns = new List<OfficeRichTextRun>();
            if (!string.IsNullOrEmpty(marker)) {
                richRuns.Add(CreateRichTextRun(marker!, firstRun, textBox, paragraph, colorScheme, mapping, markerRun: true));
            }

            for (int i = 0; i < runs.Count; i++) {
                PowerPointTextRun run = runs[i];
                if (!string.IsNullOrEmpty(run.Text)) {
                    richRuns.Add(CreateRichTextRun(run.Text, run, textBox, paragraph, colorScheme, mapping));
                }
            }

            return richRuns;
        }

        private static string CreateParagraphPlainText(PowerPointParagraph paragraph, string? marker) {
            string text = string.Concat(paragraph.Runs.Select(run => run.Text).Where(runText => !string.IsNullOrEmpty(runText)));
            return string.IsNullOrEmpty(marker) ? text : marker + text;
        }

        private static OfficeFontInfo ResolveParagraphFont(PowerPointTextBox textBox, PowerPointParagraph paragraph, PowerPointShapeBoundsMapping mapping) {
            PowerPointTextRun? firstRun = paragraph.Runs.FirstOrDefault(run => !string.IsNullOrEmpty(run.Text)) ?? paragraph.Runs.FirstOrDefault();
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

            return new OfficeFontInfo(firstRun?.FontName ?? textBox.FontName ?? "Calibri", mapping.MapFontSize(firstRun?.FontSize ?? textBox.FontSize ?? 18), style);
        }

        private static OfficeColor ResolveParagraphTextColor(PowerPointTextBox textBox, PowerPointParagraph paragraph, A.ColorScheme? colorScheme) {
            PowerPointTextRun? firstRun = paragraph.Runs.FirstOrDefault(run => !string.IsNullOrEmpty(run.Text)) ?? paragraph.Runs.FirstOrDefault();
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

        private static double EstimateParagraphTextHeight(string text, OfficeFontInfo font, double lineHeight, double textWidth, OfficeTextParagraphIndent indent) {
            OfficeTextMeasurer measurer = OfficeTextMeasurer.Create(font);
            Func<string?, double, double> measure = (value, size) => {
                OfficeTextMeasurementStyle measuredStyle = measurer.CreateStyle(new OfficeFontInfo(font.FamilyName, size, font.Style));
                return measurer.MeasureWidth(value, measuredStyle);
            };
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

        private static double EstimateParagraphRichTextHeight(IReadOnlyList<OfficeRichTextRun> runs, double maxFontSize, double lineHeight, double textWidth, OfficeTextParagraphIndent indent) {
            OfficeTextMeasurer measurer = OfficeTextMeasurer.Create();
            Func<string?, double, string?, double> measure = (value, size, family) => {
                OfficeTextMeasurementStyle measuredStyle = measurer.CreateStyle(new OfficeFontInfo(family, size));
                return measurer.MeasureWidth(value, measuredStyle);
            };
            OfficeRichTextBlockLayout layout = OfficeTextLayoutEngine.LayoutRichTextBlock(
                runs,
                textWidth,
                double.MaxValue,
                Math.Max(1D, lineHeight / Math.Max(1D, maxFontSize)),
                measure,
                wrap: true,
                shrinkToFit: false,
                minimumFontSize: Math.Min(6D, maxFontSize),
                paragraphIndent: indent);
            return Math.Max(lineHeight, layout.Height);
        }

        private static double ResolveTextBoxVerticalOffset(A.TextAnchoringTypeValues? alignment, double textHeight, double flowHeight) {
            double extraHeight = Math.Max(0D, textHeight - flowHeight);
            if (alignment == A.TextAnchoringTypeValues.Center) {
                return extraHeight / 2D;
            }

            if (alignment == A.TextAnchoringTypeValues.Bottom) {
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
