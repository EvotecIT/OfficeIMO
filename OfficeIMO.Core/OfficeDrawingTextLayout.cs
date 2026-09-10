using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>Shared native drawing text measurement used by SVG, raster and PDF adapters.</summary>
internal static class OfficeDrawingTextLayout {
    // Paragraph advance may be smaller than the final glyph box with condensed leading.
    // Keep these extents separate so consumers can diagnose paint beyond a real frame.
    internal static double PaintedHeight(OfficeTextBlockLayout layout) {
        double bottom = 0D;
        for (int i = 0; i < layout.Lines.Count; i++) {
            if (!string.IsNullOrWhiteSpace(layout.Lines[i].Text))
                bottom = Math.Max(bottom, i * layout.LineHeight + Math.Max(0D, (layout.LineHeight - layout.FontSize) / 2D) + layout.FontSize);
        }
        return bottom;
    }

    internal static double PaintedHeight(OfficeRichTextBlockLayout layout) {
        double top = 0D, bottom = 0D;
        foreach (OfficeRichTextLine line in layout.Lines) {
            double height = line.LineHeight > 0D ? line.LineHeight : layout.LineHeight;
            bool hasText = false;
            foreach (OfficeRichTextSegment segment in line.Segments)
                hasText |= !string.IsNullOrWhiteSpace(segment.Text);
            if (hasText) {
                OfficeTextLayoutEngine.ResolveRichTextVerticalExtents(line.Segments, out _, out double contentBottom);
                bottom = Math.Max(bottom, OfficeTextBlockRenderer.ResolveRichTextRenderBaseline(line, top, height, true) + contentBottom);
            }
            top += height;
        }
        return bottom;
    }

    internal static double ResolveLineHeightFactor(double? lineHeight, double fontSize) =>
        lineHeight.HasValue && lineHeight.Value > 0D && fontSize > 0D
            ? lineHeight.Value / fontSize : 1.2D;

    internal static OfficeRasterCanvas CreateMetrics(OfficeFontFaceCollection? fonts) =>
        new OfficeRasterCanvas(new OfficeRasterImage(1, 1), fonts: fonts);

    internal static OfficeRasterCanvas CreateMetrics(OfficeDrawing drawing, System.Threading.CancellationToken cancellationToken = default) =>
        new OfficeRasterCanvas(new OfficeRasterImage(1, 1), null, drawing.Fonts,
            drawing.TextShapingProvider, drawing.TextShapingLanguage, cancellationToken: cancellationToken);

    internal static OfficeRichTextBlockLayout Create(
        OfficeDrawingRichText text, double width, double height,
        Func<string?, double, string?, OfficeFontStyle, double> measure, double scale = 1D) {
        var runs = new List<OfficeRichTextRun>(text.Runs.Count);
        double maxFontSize = 0D;
        foreach (OfficeRichTextRun run in text.Runs) {
            maxFontSize = Math.Max(maxFontSize, run.FontSize * scale);
            runs.Add(new OfficeRichTextRun(run.Text, run.FontSize * scale, run.Color,
                run.Bold, run.Italic, run.Underline, run.FontFamily, run.Strikethrough,
                run.BackgroundColor, run.UnderlineStyle, run.StrikethroughStyle, run.Baseline) {
                LinkUri = run.LinkUri
            });
        }
        if (maxFontSize <= 0D) maxFontSize = 10D * scale;
        double factor = ResolveLineHeightFactor(text.LineHeight * scale, maxFontSize);
        return OfficeTextLayoutEngine.LayoutStyledRichTextBlock(runs, width, height,
            factor, measure, text.WrapText, text.ShrinkToFit, Math.Min(6D * scale, maxFontSize),
            paragraphIndent: text.ParagraphIndent.Scale(scale), shrinkToHeight: true);
    }
}
