using System;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    private OfficePreparedTextShaping? _preparedShaping;
    private IOfficeFontProgram? _preparedMetricFont;
    private OfficeTrueTypeFont.ShapedTextRun? _preparedMetricRun;
    private double _preparedMetricScale;

    // State belongs to this canvas and one text element; there is no static/thread ambient state.
    internal PreparedShapingScope PushPreparedShaping(OfficePreparedTextShaping? prepared) {
        if (ReferenceEquals(_preparedShaping, prepared)) return default;
        var scope = new PreparedShapingScope(this, _preparedShaping);
        SetPreparedShaping(prepared);
        return scope;
    }

    private void SetPreparedShaping(OfficePreparedTextShaping? prepared) {
        if (ReferenceEquals(_preparedShaping, prepared)) return;
        _preparedShaping = prepared;
        _preparedMetricFont = null;
        _preparedMetricRun = null;
        _textMeasurementCache?.Clear();
        _textLineInkBoundsCache?.Clear();
        _textPaintBoundsCache?.Clear();
    }

    private IOfficeFontProgram ResolvePreparedFont() {
        if (_preparedMetricFont == null || _preparedMetricScale != FontMetricScale) {
            _preparedMetricFont = ResolveMetricScale(_preparedShaping!.Program)!;
            _preparedMetricRun = _preparedMetricFont is OfficeTrueTypeFont trueType
                ? _preparedShaping.WithMetricFont(trueType) : null;
            _preparedMetricScale = FontMetricScale;
        }
        return _preparedMetricFont;
    }

    private bool IsPreparedText(string text, IOfficeFontProgram font) =>
        _preparedShaping != null && string.Equals(text, _preparedShaping.Text, StringComparison.Ordinal) &&
        ReferenceEquals(font, ResolvePreparedFont());

    internal struct PreparedShapingScope : IDisposable {
        private OfficeRasterCanvas? _canvas;
        private readonly OfficePreparedTextShaping? _previous;
        internal PreparedShapingScope(OfficeRasterCanvas canvas, OfficePreparedTextShaping? previous) {
            _canvas = canvas; _previous = previous;
        }
        public void Dispose() {
            if (_canvas == null) return;
            _canvas.SetPreparedShaping(_previous);
            _canvas = null;
        }
    }
}
