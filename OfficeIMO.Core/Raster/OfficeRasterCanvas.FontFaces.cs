using System;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    private OfficeFontFaceDescriptor? _textFace;

    /// <summary>Measures text using exact face attributes and the same fallback plan as painting.</summary>
    public double MeasureText(string? text, OfficeFontInfo font) {
        using (PushTextFace(font.Face)) return MeasureText(text, font.Size, font.FamilyName, font.Style);
    }

    /// <summary>Draws text using the font's numeric face attributes and text decorations.</summary>
    public void DrawText(string? text, double x, double y, double width, double height, OfficeColor color,
        OfficeFontInfo font, OfficeTextAlignment alignment = OfficeTextAlignment.Left) {
        using (PushTextFace(font.Face)) DrawText(text, x, y, width, height, color, font.Size, alignment, font.Style, font.FamilyName);
    }

    internal OfficeFontFaceDescriptor RequestedTextFace(OfficeFontStyle style) => _textFace ?? OfficeFontFaceDescriptor.FromStyle(style);

    // One renderer operation may measure, split, transform and paint a text block through
    // legacy convenience methods. Keep its exact request intact throughout those calls.
    internal IDisposable PushTextFace(OfficeFontFaceDescriptor face) {
        var scope = new FontFaceScope(this, _textFace);
        _textFace = face;
        return scope;
    }

    private sealed class FontFaceScope : IDisposable {
        private OfficeRasterCanvas? _canvas;
        private readonly OfficeFontFaceDescriptor? _previous;
        internal FontFaceScope(OfficeRasterCanvas canvas, OfficeFontFaceDescriptor? previous) { _canvas = canvas; _previous = previous; }
        public void Dispose() { if (_canvas != null) { _canvas._textFace = _previous; _canvas = null; } }
    }
}
