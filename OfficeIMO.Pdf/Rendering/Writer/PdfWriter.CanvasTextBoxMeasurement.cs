using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>Canvas composition and conservative paint bounds in coordinates relative to the box's top left.</summary>
    internal sealed record CanvasTextBoxMeasurement(double ContentHeight, bool TextFits, double TextInkTop,
        double PaintLeft, double PaintTop, double PaintWidth, double PaintHeight);

    /// <summary>Uses the renderer's resolved fonts, line placement, insets and shape stroke geometry.</summary>
    internal static CanvasTextBoxMeasurement? MeasureCanvasTextBox(IReadOnlyList<PdfTextRun> runs,
        double width, double height, PdfCanvasTextBoxStyle style, PdfCreateOptions options) {
        double textHeight = height - style.EffectivePaddingTop - style.EffectivePaddingBottom;
        if (!IsFiniteCanvasMeasurement(width) || !IsFiniteCanvasMeasurement(height) ||
            width <= style.EffectivePaddingLeft + style.EffectivePaddingRight || textHeight <= 0D) return null;
        PdfStandardFont baseFont = ChooseNormal(style.Font ?? options.DefaultFont);
        if (NormalizeFallbackRuns(runs, baseFont, options)
            .Any(run => !CanWriteRunWithSelectedFont(run, baseFont, options))) return null;
        var prepared = PrepareCanvasTextBoxText(runs, width, style, options);
        if (!IsFiniteCanvasMeasurement(prepared.ContentHeight)) return null;

        double left = double.PositiveInfinity, top = double.PositiveInfinity;
        double right = double.NegativeInfinity, bottom = double.NegativeInfinity;
        bool fits = prepared.ContentHeight <= textHeight + .01D;
        double inkTop = double.PositiveInfinity, inkBottom = double.NegativeInfinity;
        double lineBaseline = -GetAscenderForOptions(prepared.BaseFont, prepared.FontSize, options)
            - GetCanvasTextBoxVerticalOffset(style.VerticalAlign, textHeight, prepared.ContentHeight);
        for (int index = 0; index < prepared.Lines.Count; index++) {
            List<RichSeg> line = prepared.Lines[index];
            if (line.Any(segment => !string.IsNullOrWhiteSpace(segment.Text) || segment.InlineElement != null ||
                    segment.Text.Length > 0 && (segment.BackgroundColor.HasValue || segment.Underline || segment.Strike))) {
                double baseline = AdjustRichLineBaseline(lineBaseline, line, options,
                    prepared.FontSize, prepared.BaseFont);
                GetRichLineInkMetrics(line, options, out double ascent, out double descent);
                inkTop = Math.Min(inkTop, -(baseline + ascent));
                inkBottom = Math.Max(inkBottom, -(baseline - descent));
            }
            lineBaseline -= GetRichLineHeight(prepared.LineHeights, index, prepared.Leading);
        }
        bool unmeasuredDecoration = runs.Any(run => run.Underline || run.Strike || run.BackgroundColor.HasValue);
        if (unmeasuredDecoration) Include(style.EffectivePaddingLeft, style.EffectivePaddingTop, prepared.TextWidth, textHeight);
        if (!double.IsPositiveInfinity(inkTop)) {
            if (!IsFiniteCanvasMeasurement(inkTop) || !IsFiniteCanvasMeasurement(inkBottom)) return null;
            // Exact spacing can fit its advances while clipping a larger run's ink.
            fits &= inkTop >= -.01D && inkBottom <= textHeight + .01D;
            double paintTop = unmeasuredDecoration ? 0D : Math.Max(0D, inkTop);
            double paintBottom = unmeasuredDecoration ? textHeight : Math.Min(textHeight, inkBottom);
            Include(style.EffectivePaddingLeft, style.EffectivePaddingTop + paintTop,
                prepared.TextWidth, paintBottom - paintTop);
        }
        OfficeShape? shape = CreateCanvasTextBoxShape(width, height, style);
        if (shape != null && TryGetShapePaintBounds(shape, 0D, 0D, null, "Canvas text box", "shape",
                out double shapeLeft, out double shapeBottom, out double shapeWidth, out double shapeHeight,
                conservativeWhenUnmeasured: true))
            Include(shapeLeft, height - shapeBottom - shapeHeight, shapeWidth, shapeHeight);
        double textInkTop = double.IsPositiveInfinity(inkTop) ? 0D : inkTop;
        if (double.IsPositiveInfinity(left)) return new(prepared.ContentHeight, fits, textInkTop, 0D, 0D, 0D, 0D);
        if (!IsFiniteCanvasMeasurement(left) || !IsFiniteCanvasMeasurement(top) ||
            !IsFiniteCanvasMeasurement(right) || !IsFiniteCanvasMeasurement(bottom)) return null;
        return new(prepared.ContentHeight, fits, textInkTop, left, top, right - left, bottom - top);

        void Include(double x, double y, double w, double h) {
            if (w <= 0D || h <= 0D) return;
            left = Math.Min(left, x); top = Math.Min(top, y);
            right = Math.Max(right, x + w); bottom = Math.Max(bottom, y + h);
        }
    }

    private static bool IsFiniteCanvasMeasurement(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    private static double GetCanvasTextBoxVerticalOffset(PdfVerticalAlign align, double boxHeight, double contentHeight) {
        double unusedHeight = Math.Max(0D, boxHeight - contentHeight);
        return align switch {
            PdfVerticalAlign.Middle => unusedHeight / 2D,
            PdfVerticalAlign.Bottom => unusedHeight,
            _ => 0D
        };
    }

    // Keep RichLine instances: they carry each composed line's resolved baseline.
    private static (double FontSize, double Leading, PdfStandardFont BaseFont, double TextWidth,
        List<List<RichSeg>> Lines, List<double> LineHeights, double ContentHeight)
        PrepareCanvasTextBoxText(IReadOnlyList<PdfTextRun> runs, double width,
            PdfCanvasTextBoxStyle style, PdfCreateOptions options) {
        double textWidth = width - style.EffectivePaddingLeft - style.EffectivePaddingRight;
        double size = style.FontSize ?? options.DefaultFontSize;
        double leading = style.LineHeight ?? size * 1.2D;
        PdfStandardFont baseFont = ChooseNormal(style.Font ?? options.DefaultFont);
        var wrap = WrapRichRunsWithSpacing(runs, textWidth, size, baseFont, leading,
            null, DefaultParagraphTabStopWidth, options, style.LineSpacing);
        return (size, leading, baseFont, textWidth, wrap.Lines, wrap.LineHeights,
            MeasureRichLinesHeight(wrap.LineHeights, wrap.Lines.Count, leading));
    }

    private static OfficeShape? CreateCanvasTextBoxShape(double width, double height, PdfCanvasTextBoxStyle style) {
        if (!style.Background.HasValue && !(style.BorderColor.HasValue && style.BorderWidth > 0D)) return null;
        OfficeShape shape = style.CornerRadius > 0D
            ? OfficeShape.RoundedRectangle(width, height, style.CornerRadius) : OfficeShape.Rectangle(width, height);
        shape.FillColor = style.Background?.ToOfficeColor();
        shape.FillOpacity = style.Background.HasValue ? style.BackgroundOpacity : null;
        shape.StrokeColor = style.BorderWidth > 0D ? style.BorderColor?.ToOfficeColor() : null;
        shape.StrokeWidth = style.BorderWidth;
        shape.StrokeDashStyle = style.BorderDashStyle;
        shape.StrokeLineCap = style.BorderLineCap;
        shape.StrokeLineJoin = style.BorderLineJoin;
        return shape;
    }
}
