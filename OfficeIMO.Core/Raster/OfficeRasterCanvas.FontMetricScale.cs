using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    private double _fontMetricScale = 1D;
    internal double FontMetricScale {
        get => _fontMetricScale;
        set {
            if (value == _fontMetricScale) return;
            if (value <= 0D || double.IsNaN(value) || double.IsInfinity(value)) throw new ArgumentOutOfRangeException(nameof(value));
            _fontMetricScale = value;
            _scaledMetricFonts?.Clear();
        }
    }

    internal IDisposable PushFontMetricScale(double scale) {
        var scope = new FontMetricScaleScope(this, FontMetricScale);
        FontMetricScale = scale;
        return scope;
    }

    private sealed class FontMetricScaleScope : IDisposable {
        private OfficeRasterCanvas? _canvas;
        private readonly double _previous;
        internal FontMetricScaleScope(OfficeRasterCanvas canvas, double previous) { _canvas = canvas; _previous = previous; }
        public void Dispose() { if (_canvas != null) { _canvas.FontMetricScale = _previous; _canvas = null; } }
    }
    private Dictionary<OfficeTrueTypeFont, OfficeTrueTypeFont>? _scaledMetricFonts;

    private IOfficeFontProgram? ResolveMetricScale(IOfficeFontProgram? font) {
        if (font is not OfficeTrueTypeFont trueType) return font;
        if (FontMetricScale == 1D) return trueType;
        var cache = _scaledMetricFonts ??= new Dictionary<OfficeTrueTypeFont, OfficeTrueTypeFont>();
        if (cache.TryGetValue(trueType, out OfficeTrueTypeFont? instance)) return instance;
        if (cache.Count >= 1024) cache.Clear();
        instance = trueType.ForOutputScale(FontMetricScale);
        cache[trueType] = instance;
        return instance;
    }
}
