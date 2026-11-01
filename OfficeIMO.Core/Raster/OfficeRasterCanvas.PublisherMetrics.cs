using System;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    internal OfficePublisherFontMetrics? MeasurePublisherFont(string? text, double size, string? family, OfficeFontStyle style) {
        _cancellationToken.ThrowIfCancellationRequested();
        if (_fonts != null && !string.IsNullOrEmpty(text)) {
            var runs = _fonts.PlanFallbackRuns(text, family, RequestedTextFace(style));
            if (ShouldUseFallbackRuns(runs, family)) {
                double ascent = 0D, descent = 0D, baseline = 0D;
                bool resolved = false;
                foreach (OfficeFontFallbackRun run in runs) {
                    OfficePublisherFontMetrics? selected = MeasurePublisherFont(run.Text, size, run.FamilyRequest, style);
                    OfficePublisherFontMetrics partMetrics = selected ?? new OfficePublisherFontMetrics(size * .84D, size * .16D, size * .84D);
                    ascent = Math.Max(ascent, partMetrics.Ascent); descent = Math.Max(descent, partMetrics.Descent);
                    baseline = Math.Max(baseline, partMetrics.Baseline); resolved |= selected.HasValue;
                }
                return resolved ? new OfficePublisherFontMetrics(ascent, descent, baseline) : null;
            }
        }
        IOfficeFontProgram? font = ResolveTextFont(text, family, style, size);
        OfficePublisherFontMetrics? metrics = font == null ? null : OfficePublisherFontMetrics.Resolve(font, size / FontMetricScale);
        return metrics is { } value ? new OfficePublisherFontMetrics(value.Ascent * FontMetricScale,
            value.Descent * FontMetricScale, value.Baseline * FontMetricScale) : null;
    }

    internal OfficeDrawingTextMetrics CreateDrawingTextMetrics() => new OfficeDrawingTextMetrics(MeasureText, MeasureTextPaintBounds,
        (text, size, family, style) => MeasureTextLineHorizontalPaintBounds(text ?? string.Empty, size, family, style),
        null, (segment, size) => OfficeDrawingTextLayout.MeasureSegmentPaintBounds(segment, size, MeasureTextPaintBounds),
        MeasurePublisherFont);
}
