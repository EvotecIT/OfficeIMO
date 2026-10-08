namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    // A retained child drawing can override text shaping without changing the
    // parent or its siblings. Keep the same target, font scope, clip and diagnostics.
    internal OfficeRasterCanvas WithDrawingTextProfile(OfficeDrawing drawing) =>
        WithTextShapingProfile(drawing.TextShapingProvider, drawing.TextShapingLanguage);

    internal OfficeRasterCanvas WithTextShapingProfile(IOfficeTextShapingProvider? provider, string? language) {
        provider ??= _textShapingProvider;
        language = NormalizeTextShapingLanguage(language) ?? _textShapingLanguage;
        if (ReferenceEquals(provider, _textShapingProvider) && language == _textShapingLanguage) return this;
        OfficeRasterCanvas child = _image != null
            ? new OfficeRasterCanvas(_image, _font, _fonts, provider, language, _diagnosticSink, _diagnosticSource, _cancellationToken)
            : new OfficeRasterCanvas(_target!, _font, _fonts, provider, language, _diagnosticSink, _diagnosticSource, _cancellationToken);
        child.FontMetricScale = FontMetricScale;
        child.SetCoordinateScale(CoordinateScaleX, CoordinateScaleY);
        child.ShareTransformedTextBudget(_transformedTextBudget);
        child._clipRegion = _clipRegion;
        child.PreservePaintedGlyphOrder = PreservePaintedGlyphOrder;
        return child;
    }
}
