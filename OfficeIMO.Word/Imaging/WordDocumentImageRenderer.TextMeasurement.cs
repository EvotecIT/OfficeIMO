using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using OfficeIMO.Drawing;

namespace OfficeIMO.Word;

internal static partial class WordDocumentImageRenderer {
    // Core's plain text layout rounds each line up before deciding how many fit.
    // Allocate that same height: two 11pt lines need 28 units, rather than 27.5.
    private static double ResolvePlainTextLineHeight(OfficeFontInfo font) =>
        Math.Ceiling(Math.Max(font.Size * 1.25D, 12D));

    // Secondary layout passes must use the same font faces and shaping profile as the painted page.
    private static OfficeDrawing CreateMeasurementDrawing(OfficeDrawing source, double width) {
        var drawing = new OfficeDrawing(Math.Max(1D, width), double.MaxValue) {
            TextShapingProvider = source.TextShapingProvider,
            TextShapingLanguage = source.TextShapingLanguage
        };
        drawing.Fonts.AddRange(source.Fonts);
        return drawing;
    }

    private static double EstimateTextHeight(string text, OfficeFontInfo font, double contentWidth,
        double lineHeight, OfficeRasterCanvas metrics, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        OfficeTextBlockLayout layout = OfficeTextLayoutEngine.LayoutTextBlock(
            text, font.Size, Math.Max(1D, contentWidth), double.MaxValue,
            lineHeight / font.Size, Math.Min(6D, font.Size),
            (value, size) => {
                cancellationToken.ThrowIfCancellationRequested();
                return metrics.MeasureText(value, size, font.FamilyName, font.Style);
            }, wrap: true, forceSingleLine: false, shrinkToFit: false,
            overflowBehavior: OfficeTextOverflowBehavior.Clip);
        return Math.Max(lineHeight, layout.Height);
    }

    private static double ResolveRichTextFrameLineHeight(IReadOnlyList<OfficeRichTextRun> runs) =>
        Math.Max(runs.Max(run => run.FontSize) * 1.25D, 12D);

    private static double EstimateRichTextFrameHeight(IReadOnlyList<OfficeRichTextRun> runs, double contentWidth,
        OfficeRasterCanvas metrics, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var text = new OfficeDrawingRichText(runs, 0D, 0D, Math.Max(1D, contentWidth), double.MaxValue,
            lineHeight: ResolveRichTextFrameLineHeight(runs), wrapText: true);
        // Use the painter's shared rich layout, including per-run styles and line heights.
        OfficeRichTextBlockLayout layout = OfficeDrawingTextLayout.Create(text, text.Width, text.Height,
            (value, size, family, style) => {
                cancellationToken.ThrowIfCancellationRequested();
                return metrics.MeasureText(value, size, family, style);
            }, measurePaint: (value, size, family, style) => {
                cancellationToken.ThrowIfCancellationRequested();
                return metrics.MeasureTextPaintBounds(value, size, family, style);
            });
        return layout.Height;
    }

    private static Func<string?, double, string?, OfficeFontStyle, double> CreateRichTextMeasure(WordImageFlowContext context) {
        OfficeRasterCanvas metrics = context.TextMetrics;
        return (value, size, family, style) => {
            context.ThrowIfCancellationRequested();
            return metrics.MeasureText(value, size, family, style);
        };
    }
}
