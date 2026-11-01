using System;

namespace OfficeIMO.Drawing;

// The old font-only callback remains the default. Drawing adapters additionally
// provide feature-aware measurement so each run's policy follows its actual face.
internal sealed class OfficeRichTextMeasurement {
    private readonly Func<string?, double, string?, OfficeFontStyle, double> _measure;
    private readonly Func<string?, double, string?, OfficeFontStyle, OfficeTextFeatureSettings?, double>? _measureFeatures;
    internal bool HasFeatureMeasurement => _measureFeatures != null;

    internal OfficeRichTextMeasurement(Func<string?, double, string?, OfficeFontStyle, double> measure,
        Func<string?, double, string?, OfficeFontStyle, OfficeTextFeatureSettings?, double>? measureFeatures = null) {
        _measure = measure ?? throw new ArgumentNullException(nameof(measure));
        _measureFeatures = measureFeatures;
    }

    internal double Measure(string? text, double size, string? family, OfficeFontStyle style) =>
        Measure(text, size, family, style, null);
    internal double Measure(string? text, OfficeRichTextRun run) =>
        Measure(text, run.EffectiveFontSize, run.FontFamily, run.FontStyle, run.KerningFeatures);
    internal double Measure(string? text, OfficeRichTextSegment segment) =>
        Measure(text, OfficeTextBlockRenderer.ResolveRichTextRenderedFontSize(segment), segment.FontFamily, segment.FontStyle, segment.KerningFeatures);

    private double Measure(string? text, double size, string? family, OfficeFontStyle style, OfficeTextFeatureSettings? features) {
        if (string.IsNullOrEmpty(text)) return 0D;
        if (size <= 0D || double.IsNaN(size) || double.IsInfinity(size)) size = 1D;
        return Math.Max(0D, features != null && _measureFeatures != null
            ? _measureFeatures(text, size, family, style, features) : _measure(text, size, family, style));
    }
}

internal static class OfficeTextKerning {
    private static readonly OfficeTextFeatureSettings Enabled = OfficeTextFeatureSettings.Default.With("kern", 1);
    private static readonly OfficeTextFeatureSettings Disabled = OfficeTextFeatureSettings.Default.With("kern", 0);

    // Drawing scale transforms both point values. Font fitting changes size only;
    // superscript/subscript use the base run size before glyph-size reduction.
    internal static OfficeTextFeatureSettings? Features(double? threshold, double size) =>
        !threshold.HasValue ? null : threshold.Value > 0D && size >= threshold.Value ? Enabled : Disabled;
}
