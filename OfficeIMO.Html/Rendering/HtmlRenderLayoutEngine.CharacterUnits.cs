namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    // A ch is the used zero-glyph advance, without letter or word spacing.
    // Select the same scoped face and shaping features as an ordinary text run.
    private double MeasureCharacterAdvance(HtmlRenderBoxStyle style) {
        if (IsVerticalWritingMode(style.WritingMode) && style.TextOrientation == "upright") return style.Font.Size;
        var fallback = _fonts.PlanFallbackRuns("0", style.Font.FamilyName, style.Font.Face).FirstOrDefault();
        HtmlRenderBoxStyle measurementStyle = style.Clone();
        measurementStyle.Font = style.Font.WithFamilyName(fallback?.FamilyName ?? style.Font.FamilyName);
        measurementStyle.BaselineScale = 1D;
        double advance = TryMeasureWithConfiguredProvider("0", measurementStyle, out double shaped)
            ? shaped : MeasureText("0", measurementStyle.Font);
        return double.IsNaN(advance) || double.IsInfinity(advance) || advance < 0D ? style.Font.Size * 0.5D : advance;
    }

    private bool TryResolveLength(string? value, double reference, HtmlRenderBoxStyle style, out double result) =>
        HtmlRenderCssValues.TryLength(value, reference, style.Font.Size, _options.DefaultFontSize,
            _options.Mode == HtmlRenderMode.Paged ? _activePageGeometry.Width : _options.ViewportWidth,
            _options.Mode == HtmlRenderMode.Paged ? _activePageGeometry.Height : _options.ViewportHeight ?? 1056D,
            style.ContainerUnitWidth ?? double.NaN, style.ContainerUnitHeight ?? double.NaN, out result, style.CharacterAdvance);
}
