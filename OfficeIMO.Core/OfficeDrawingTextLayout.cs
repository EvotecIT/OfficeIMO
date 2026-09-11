using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>Shared native drawing text measurement used by SVG, raster and PDF adapters.</summary>
internal static class OfficeDrawingTextLayout {
    internal static OfficeTextBlockLayout IncludePaintedHeight(OfficeTextBlockLayout layout, double availableHeight,
        Func<string?, double, OfficeTextPaintBounds>? measurePaint = null) {
        OfficeTextPaintBounds bounds = PaintedBounds(layout, measurePaint);
        double offset = bounds.Top < -1e-9D ? -bounds.Top : 0D;
        double height = Math.Max(layout.Height, bounds.Bottom) + offset;
        if (height - layout.Height < 1e-9D) height = layout.Height;
        return new OfficeTextBlockLayout(layout.Lines, layout.FontSize, layout.LineHeight, layout.Width,
            height, layout.Clipped || height > availableHeight + .01D) { ContentOffsetY = offset };
    }

    internal static OfficeRichTextBlockLayout IncludePaintedHeight(OfficeRichTextBlockLayout layout, double availableHeight,
        Func<string?, double, string?, OfficeFontStyle, OfficeTextPaintBounds>? measurePaint = null) {
        OfficeTextPaintBounds bounds = PaintedBounds(layout, measurePaint);
        double offset = bounds.Top < -1e-9D ? -bounds.Top : 0D;
        double height = Math.Max(layout.Height, bounds.Bottom) + offset;
        if (height - layout.Height < 1e-9D) height = layout.Height;
        return new OfficeRichTextBlockLayout(layout.Lines, layout.LineHeight, layout.Width,
            height, layout.Clipped || height > availableHeight + .01D) { ContentOffsetY = offset };
    }

    internal static double RequiredFrameHeight(OfficeTextBlockLayout layout,
        Func<string?, double, OfficeTextPaintBounds>? measurePaint = null) {
        OfficeTextPaintBounds bounds = PaintedBounds(layout, measurePaint);
        return Math.Max(layout.Height, bounds.Bottom) + Math.Max(0D, -bounds.Top);
    }

    internal static double RequiredFrameHeight(OfficeRichTextBlockLayout layout,
        Func<string?, double, string?, OfficeFontStyle, OfficeTextPaintBounds>? measurePaint = null) {
        OfficeTextPaintBounds bounds = PaintedBounds(layout, measurePaint);
        return Math.Max(layout.Height, bounds.Bottom) + Math.Max(0D, -bounds.Top);
    }

    internal static double PaintedHeight(OfficeTextBlockLayout layout,
        Func<string?, double, OfficeTextPaintBounds>? measurePaint = null) {
        OfficeTextPaintBounds bounds = PaintedBounds(layout, measurePaint);
        return bounds.Bottom - Math.Min(0D, bounds.Top);
    }

    internal static double PaintedHeight(OfficeRichTextBlockLayout layout,
        Func<string?, double, string?, OfficeFontStyle, OfficeTextPaintBounds>? measurePaint = null) {
        OfficeTextPaintBounds bounds = PaintedBounds(layout, measurePaint);
        return bounds.Bottom - Math.Min(0D, bounds.Top);
    }

    private static OfficeTextPaintBounds PaintedBounds(OfficeTextBlockLayout layout,
        Func<string?, double, OfficeTextPaintBounds>? measurePaint) {
        double top = 0D, bottom = 0D;
        for (int i = 0; i < layout.Lines.Count; i++) {
            string text = layout.Lines[i].Text;
            if (string.IsNullOrWhiteSpace(text)) continue;
            double size = layout.FontSize;
            double baseline = i * layout.LineHeight + Math.Max(0D, (layout.LineHeight - size) / 2D) + size * .84D;
            OfficeTextPaintBounds bounds = measurePaint?.Invoke(text, size) ?? new OfficeTextPaintBounds(-size * .84D, size * .16D);
            top = Math.Min(top, baseline + bounds.Top); bottom = Math.Max(bottom, baseline + bounds.Bottom);
        }
        return new OfficeTextPaintBounds(top, bottom);
    }

    private static OfficeTextPaintBounds PaintedBounds(OfficeRichTextBlockLayout layout,
        Func<string?, double, string?, OfficeFontStyle, OfficeTextPaintBounds>? measurePaint) {
        double lineTop = 0D, top = 0D, bottom = 0D;
        foreach (OfficeRichTextLine line in layout.Lines) {
            double height = OfficeTextBlockRenderer.ResolveRichTextRenderLineHeight(line, layout.LineHeight);
            double baseline = OfficeTextBlockRenderer.ResolveRichTextRenderBaseline(line, lineTop, height, true);
            foreach (OfficeRichTextSegment segment in line.Segments) {
                if (string.IsNullOrWhiteSpace(segment.Text)) continue;
                double size = OfficeTextBlockRenderer.ResolveRichTextRenderedFontSize(segment);
                double renderedBaseline = OfficeTextBlockRenderer.ResolveRichTextRenderedBaseline(segment, baseline);
                OfficeTextPaintBounds bounds = measurePaint?.Invoke(segment.Text, size, segment.FontFamily, segment.FontStyle)
                    ?? new OfficeTextPaintBounds(-size * .84D, size * .16D);
                top = Math.Min(top, renderedBaseline + bounds.Top); bottom = Math.Max(bottom, renderedBaseline + bounds.Bottom);
            }
            lineTop += height;
        }
        return new OfficeTextPaintBounds(top, bottom);
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
        Func<string?, double, string?, OfficeFontStyle, double> measure, double scale = 1D,
        Func<string?, double, string?, OfficeFontStyle, OfficeTextPaintBounds>? measurePaint = null) {
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
        OfficeRichTextBlockLayout layout = OfficeTextLayoutEngine.LayoutStyledRichTextBlock(runs, width, height,
            factor, measure, text.WrapText, text.ShrinkToFit, Math.Min(6D * scale, maxFontSize),
            paragraphIndent: text.ParagraphIndent.Scale(scale), shrinkToHeight: true, measurePaint: measurePaint);
        return text.ShrinkToFit ? IncludePaintedHeight(layout, height, measurePaint) : layout;
    }
}
