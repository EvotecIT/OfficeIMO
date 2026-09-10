namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    // A retained child drawing can override text shaping without changing the
    // parent or its siblings. Keep the same target, font scope, clip and diagnostics.
    internal OfficeRasterCanvas WithDrawingTextProfile(OfficeDrawing drawing) {
        IOfficeTextShapingProvider? provider = drawing.TextShapingProvider ?? _textShapingProvider;
        string? language = NormalizeTextShapingLanguage(drawing.TextShapingLanguage) ?? _textShapingLanguage;
        if (ReferenceEquals(provider, _textShapingProvider) && language == _textShapingLanguage) return this;
        OfficeRasterCanvas child = _image != null
            ? new OfficeRasterCanvas(_image, _font, _fonts, provider, language, _diagnosticSink, _diagnosticSource, _cancellationToken)
            : new OfficeRasterCanvas(_target!, _font, _fonts, provider, language, _diagnosticSink, _diagnosticSource, _cancellationToken);
        child._clipRegion = _clipRegion;
        return child;
    }
}
