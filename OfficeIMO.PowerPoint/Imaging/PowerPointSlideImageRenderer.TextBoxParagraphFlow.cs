using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointSlideImageRenderer {
        private static List<PowerPointParagraphDrawing> CreateTextBoxParagraphDrawings(
            PowerPointTextBox textBox,
            IReadOnlyList<PowerPointParagraph> paragraphs,
            PowerPointParagraphNumberingState numberingState,
            double textWidth,
            PowerPointShapeBoundsMapping mapping,
            A.ColorScheme? colorScheme, Func<string?, double, string?, OfficeFontStyle, double> measure, string? defaultFontFamily,
            double fontScale = 1D, double lineSpaceReduction = 0D) {
            var results = new List<PowerPointParagraphDrawing>(paragraphs.Count);
            for (int i = 0; i < paragraphs.Count; i++) {
                PowerPointParagraph paragraph = paragraphs[i];
                string? marker = numberingState.CreateMarker(paragraph);
                OfficeTextParagraphIndent indent = CreateParagraphIndent(paragraph, mapping);
                OfficeTextAlignment alignment = MapTextAlignment(paragraph.Alignment);
                List<OfficeRichTextRun> richRuns = CreateParagraphRichTextRuns(textBox, paragraph, marker, colorScheme, mapping, defaultFontFamily);
                if (fontScale != 1D) richRuns = ScaleTextBoxRuns(richRuns, fontScale);
                double maxFontSize = richRuns.Count == 0
                    ? ResolveParagraphFont(textBox, paragraph, mapping, defaultFontFamily).Size * fontScale
                    : richRuns.Max(run => run.FontSize);
                double lineHeight = paragraph.LineSpacingPoints.HasValue
                    ? ResolveParagraphLineHeight(paragraph, maxFontSize, mapping)
                    : Math.Max(1D, maxFontSize * Math.Max(0.01D, (paragraph.LineSpacingMultiplier ?? 1.2D) - lineSpaceReduction));
                double height, paintedHeight, measuredWidth;
                bool clipped;
                if (ShouldRenderParagraphRichText(richRuns, marker)) {
                    height = EstimateParagraphRichTextHeight(richRuns, maxFontSize, lineHeight, textWidth, indent, out paintedHeight, out measuredWidth, out clipped, measure);
                    results.Add(PowerPointParagraphDrawing.FromRichText(paragraph, richRuns, alignment, indent, lineHeight, height, paintedHeight, mapping, measuredWidth, clipped));
                } else {
                    string text = richRuns.Count == 1 ? richRuns[0].Text : CreateParagraphPlainText(paragraph, marker);
                    OfficeFontInfo font = richRuns.Count == 1 ? new OfficeFontInfo(richRuns[0].FontFamily, richRuns[0].FontSize, richRuns[0].FontStyle) : ResolveParagraphFont(textBox, paragraph, mapping, defaultFontFamily);
                    if (richRuns.Count != 1 && fontScale != 1D) font = new OfficeFontInfo(font.FamilyName, font.Size * fontScale, font.Style);
                    height = EstimateParagraphTextHeight(text, font, lineHeight, textWidth, indent, out paintedHeight, out measuredWidth, out clipped, measure);
                    OfficeColor color = richRuns.Count == 1 ? richRuns[0].Color : ResolveParagraphTextColor(textBox, paragraph, colorScheme);
                    results.Add(PowerPointParagraphDrawing.FromText(paragraph, text, font, color, alignment, indent, lineHeight, height, paintedHeight, mapping, measuredWidth, clipped));
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
            PowerPointEffectiveRunStyle effective = PowerPointEffectiveRunStyleResolver.ResolveParagraph(
                paragraph, textBox.TextBody?.ListStyle, textBox.MasterTextStyle);
            return new OfficeFontInfo(effective.FontName ?? textBox.FontName ?? ResolveTextBoxFallbackFont(textBox, defaultFontFamily),
                mapping.MapFontSize(effective.FontSizePoints ?? textBox.FontSize ?? 18),
                ResolveEffectiveFontStyle(effective, textBox.Bold, textBox.Italic));
        }

        private static OfficeFontStyle ResolveEffectiveFontStyle(PowerPointEffectiveRunStyle effective, bool fallbackBold, bool fallbackItalic) {
            OfficeFontStyle style = OfficeFontStyle.Regular;
            if (effective.Bold ?? fallbackBold) style |= OfficeFontStyle.Bold;
            if (effective.Italic ?? fallbackItalic) style |= OfficeFontStyle.Italic;
            if (effective.UnderlineStyle is { } underline && underline != PowerPointUnderlineStyle.None) style |= OfficeFontStyle.Underline;
            if (effective.StrikeStyle is { } strike && strike != PowerPointStrikeStyle.None) style |= OfficeFontStyle.Strikethrough;
            return style;
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

        private static double EstimateParagraphTextHeight(string text, OfficeFontInfo font, double lineHeight, double textWidth, OfficeTextParagraphIndent indent, out double paintedHeight, Func<string?, double, string?, OfficeFontStyle, double>? styledMeasure = null) =>
            EstimateParagraphTextHeight(text, font, lineHeight, textWidth, indent, out paintedHeight, out _, out _, styledMeasure);

        private static double EstimateParagraphTextHeight(string text, OfficeFontInfo font, double lineHeight, double textWidth, OfficeTextParagraphIndent indent,
            out double paintedHeight, out double measuredWidth, out bool clipped, Func<string?, double, string?, OfficeFontStyle, double>? styledMeasure) {
            styledMeasure ??= OfficeDrawingTextLayout.CreateMetrics(null).MeasureText;
            Func<string?, double, double> measure = (value, size) => styledMeasure(value, size, font.FamilyName, font.Style);
            OfficeTextBlockLayout layout = OfficeTextLayoutEngine.LayoutTextBlock(
                text,
                font.Size,
                textWidth,
                double.MaxValue,
                OfficeDrawingTextLayout.ResolveLineHeightFactor(lineHeight, font.Size),
                Math.Min(6D, font.Size),
                measure,
                wrap: true,
                paragraphIndent: indent);
            paintedHeight = OfficeDrawingTextLayout.PaintedHeight(layout);
            measuredWidth = layout.Width;
            clipped = layout.Clipped;
            return Math.Max(lineHeight, layout.Height);
        }

        private static double EstimateParagraphRichTextHeight(IReadOnlyList<OfficeRichTextRun> runs, double maxFontSize, double lineHeight, double textWidth, OfficeTextParagraphIndent indent, out double paintedHeight,
            Func<string?, double, string?, OfficeFontStyle, double> measure) =>
            EstimateParagraphRichTextHeight(runs, maxFontSize, lineHeight, textWidth, indent, out paintedHeight, out _, out _, measure);

        private static double EstimateParagraphRichTextHeight(IReadOnlyList<OfficeRichTextRun> runs, double maxFontSize, double lineHeight, double textWidth, OfficeTextParagraphIndent indent,
            out double paintedHeight, out double measuredWidth, out bool clipped, Func<string?, double, string?, OfficeFontStyle, double> measure) {
            OfficeRichTextBlockLayout layout = OfficeTextLayoutEngine.LayoutStyledRichTextBlock(
                runs, textWidth, double.MaxValue, OfficeDrawingTextLayout.ResolveLineHeightFactor(lineHeight, maxFontSize),
                measure, wrap: true, minimumFontSize: Math.Min(6D, maxFontSize), paragraphIndent: indent);
            paintedHeight = OfficeDrawingTextLayout.PaintedHeight(layout);
            measuredWidth = layout.Width;
            clipped = layout.Clipped;
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
                double paintedHeight,
                double spaceBefore,
                double spaceAfter,
                double measuredWidth = 0D,
                bool clipped = false) {
                Text = text;
                RichRuns = richRuns;
                Font = font;
                Color = color;
                Alignment = alignment;
                Indent = indent;
                LineHeight = lineHeight;
                Height = height;
                PaintedHeight = paintedHeight;
                SpaceBefore = spaceBefore;
                SpaceAfter = spaceAfter;
                MeasuredWidth = measuredWidth;
                Clipped = clipped;
            }

            internal string Text { get; }

            internal IReadOnlyList<OfficeRichTextRun> RichRuns { get; }

            internal bool HasContent => Text.Length > 0 || RichRuns.Any(run => run.Text.Length > 0);

            internal OfficeFontInfo Font { get; }

            internal OfficeColor Color { get; }

            internal OfficeTextAlignment Alignment { get; }

            internal OfficeTextParagraphIndent Indent { get; }

            internal double LineHeight { get; }

            internal double Height { get; }

            internal double PaintedHeight { get; }

            internal double MeasuredWidth { get; }

            internal bool Clipped { get; }

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
                double paintedHeight,
                PowerPointShapeBoundsMapping mapping, double measuredWidth, bool clipped) =>
                new PowerPointParagraphDrawing(
                    text,
                    Array.Empty<OfficeRichTextRun>(),
                    font,
                    color,
                    alignment,
                    indent,
                    lineHeight,
                    height,
                    paintedHeight,
                    Math.Max(0D, mapping.MapVerticalLength(paragraph.SpaceBeforePoints ?? 0D)),
                    Math.Max(0D, mapping.MapVerticalLength(paragraph.SpaceAfterPoints ?? 0D)),
                    measuredWidth,
                    clipped);

            internal static PowerPointParagraphDrawing FromRichText(
                PowerPointParagraph paragraph,
                IReadOnlyList<OfficeRichTextRun> richRuns,
                OfficeTextAlignment alignment,
                OfficeTextParagraphIndent indent,
                double lineHeight,
                double height,
                double paintedHeight,
                PowerPointShapeBoundsMapping mapping, double measuredWidth, bool clipped) =>
                new PowerPointParagraphDrawing(
                    string.Empty,
                    richRuns,
                    OfficeFontInfo.Default,
                    OfficeColor.Black,
                    alignment,
                    indent,
                    lineHeight,
                    height,
                    paintedHeight,
                    Math.Max(0D, mapping.MapVerticalLength(paragraph.SpaceBeforePoints ?? 0D)),
                    Math.Max(0D, mapping.MapVerticalLength(paragraph.SpaceAfterPoints ?? 0D)),
                    measuredWidth,
                    clipped);
        }
    }
}
