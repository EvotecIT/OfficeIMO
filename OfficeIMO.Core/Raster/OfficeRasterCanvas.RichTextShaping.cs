using System;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    /// <summary>Retains one actual horizontal font result for coordinated rich-text measurement and paint.</summary>
    internal OfficePreparedTextShaping? PrepareRichText(string text, double size, string family,
        OfficeFontStyle style, OfficeTextFeatureSettings? features) {
        _cancellationToken.ThrowIfCancellationRequested();
        if (text.Length == 0) return null;
        if (_fonts != null && _fonts.PlanFallbackRuns(text, family, RequestedTextFace(style)).Count != 1) return null;
        IOfficeFontProgram? font = ResolveTextFont(text, family, style, size, out OfficeFontStyle resolved);
        if (font is not IOfficeBoundedFontProgram || resolved != style
            || !TryGetShapedTextRun(text, font, features, OfficeTextDirection.LeftToRight, out var shaped)) return null;
        if (shaped.Direction is not (OfficeTextDirection.Auto or OfficeTextDirection.LeftToRight)) return null;
        try { return new OfficePreparedTextShaping(font, text, shaped, _cancellationToken); }
        catch (ArgumentException) { return null; } // Unsupported clusters/color glyphs keep the ordinary rendering path.
    }

    private OfficeTextPaintBounds? MeasurePreparedSegmentPaint(OfficeRichTextSegment segment, double size) {
        using var scope = PushPreparedShaping(segment.PreparedShaping);
        return OfficeDrawingTextLayout.MeasureSegmentPaintBounds(segment, size,
            (text, em, family, style) => MeasureTextPaintBoundsWithFeatures(text, em, family, style, segment.KerningFeatures));
    }

    private (double Left, double Right) MeasurePreparedSegmentHorizontalPaint(OfficeRichTextSegment segment, double size) {
        using var scope = PushPreparedShaping(segment.PreparedShaping);
        return MeasureTextLineHorizontalPaintBoundsWithFeatures(segment.Text, size, segment.FontFamily,
            segment.FontStyle, segment.KerningFeatures);
    }
}
