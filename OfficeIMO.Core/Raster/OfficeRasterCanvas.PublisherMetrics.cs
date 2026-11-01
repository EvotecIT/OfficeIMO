using System;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    internal OfficePublisherFontMetrics? MeasurePublisherFont(string? text, double size, string? family, OfficeFontStyle style) {
        _cancellationToken.ThrowIfCancellationRequested();
        if (_fonts != null && !string.IsNullOrEmpty(text)) {
            var runs = _fonts.PlanFallbackRuns(text, family, RequestedTextFace(style));
            if (ShouldUseFallbackRuns(runs, family)) {
                var combined = new OfficePublisherFontMetrics.Union();
                foreach (OfficeFontFallbackRun run in runs)
                    combined.Include(MeasurePublisherFont(run.Text, size, run.FamilyRequest, style), size);
                return combined.Value;
            }
        }
        IOfficeFontProgram? font = ResolveTextFont(text, family, style, size);
        OfficePublisherFontMetrics? metrics = font == null ? null : OfficePublisherFontMetrics.Resolve(font, size / FontMetricScale);
        return metrics?.Scale(FontMetricScale);
    }

    internal double MeasureTextWithFeatures(string? text, double size, string? family, OfficeFontStyle style, OfficeTextFeatureSettings? features) =>
        features == null ? MeasureText(text, size, family, style)
            : MeasurePositionedText(text, size, family, style, features, OfficeTextDirection.Auto);

    internal OfficeDrawingTextMetrics CreateDrawingTextMetrics() => new OfficeDrawingTextMetrics(MeasureText, MeasureTextPaintBounds,
        (text, size, family, style) => MeasureTextLineHorizontalPaintBounds(text ?? string.Empty, size, family, style),
        null, (segment, size) => OfficeDrawingTextLayout.MeasureSegmentPaintBounds(segment, size,
            (text, em, family, style) => MeasureTextPaintBoundsWithFeatures(text, em, family, style, segment.KerningFeatures)),
        MeasurePublisherFont, MeasureTextWithFeatures,
        (segment, size) => MeasureTextLineHorizontalPaintBoundsWithFeatures(segment.Text, size,
            segment.FontFamily, segment.FontStyle, segment.KerningFeatures));
}
