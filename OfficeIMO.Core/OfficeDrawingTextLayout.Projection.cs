using System;

namespace OfficeIMO.Drawing;

internal static partial class OfficeDrawingTextLayout {
    /// <summary>
    /// Measures drawing projection loss at the authored content width, allowing
    /// only complete unwrapped paragraph bodies to overhang horizontally.
    /// </summary>
    internal static OfficeRichTextBlockLayout CreateForProjection(OfficeDrawingRichText text, double width, double height,
        Func<string?, double, string?, OfficeFontStyle, double> measure, double scale = 1D,
        Func<string?, double, string?, OfficeFontStyle, OfficeTextPaintBounds>? measurePaint = null) =>
        CreateForProjectionCore(text, width, height, measure, scale, measurePaint, null);

    internal static OfficeRichTextBlockLayout CreateForProjectionWithMetrics(OfficeDrawingRichText text, double width, double height,
        OfficeDrawingTextMetrics metrics, double scale = 1D) =>
        CreateWithHorizontalPaintInsets(text, width, height, metrics, scale, projection: true);

    private static OfficeRichTextBlockLayout CreateForProjectionCore(OfficeDrawingRichText text, double width, double height,
        Func<string?, double, string?, OfficeFontStyle, double> measure, double scale,
        Func<string?, double, string?, OfficeFontStyle, OfficeTextPaintBounds>? measurePaint,
        Func<OfficeRichTextSegment, double, OfficeTextPaintBounds?>? measureSegmentPaint) {
        if (text.WrapText || text.ShrinkToFit || text.Paragraphs.Count == 0)
            return CreateCore(text, width, height, measure, scale, measurePaint, measureSegmentPaint);
        return CreateParagraphs(text, width, height, measure, scale, measurePaint,
            reportUnwrappedWidthOverflow: false, measureSegmentPaint: measureSegmentPaint);
    }
}
