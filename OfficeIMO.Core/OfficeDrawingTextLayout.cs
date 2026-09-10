using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>Shared native drawing text measurement used by SVG, raster and PDF adapters.</summary>
internal static class OfficeDrawingTextLayout {
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
        double factor = text.LineHeight.HasValue && text.LineHeight.Value > 0D
            ? Math.Max(1D, text.LineHeight.Value * scale / maxFontSize) : 1.2D;
        return OfficeTextLayoutEngine.LayoutStyledRichTextBlock(runs, width, height,
            factor, measure, text.WrapText, text.ShrinkToFit, Math.Min(6D * scale, maxFontSize),
            paragraphIndent: text.ParagraphIndent.Scale(scale), shrinkToHeight: true);
    }
}
