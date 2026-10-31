namespace OfficeIMO.Pdf;

internal sealed partial class ContentStreamBuilder {
    private bool _syntheticBold, _syntheticWeightScope;
    private int _textObjectStart;
    private OfficeColor _paintFillColor = OfficeColor.Black, _textEntryFillColor = OfficeColor.Black;

    internal static double SyntheticBoldStrokeWidth(double size) => size / 24D;

    private void ApplySyntheticWeight(double size, bool bold) {
        if (bold) {
            if (!_inText) throw new InvalidOperationException("Synthetic font weight requires a text object.");
            if (!_syntheticWeightScope) {
                // The need for synthesis may first appear halfway through a rich
                // paragraph. Scope its entire text object, including already written
                // runs, so stroke width, join and colour cannot leak into later paint.
                _sb.Insert(_textObjectStart, "q\n");
                if (_showStart >= _textObjectStart) _showStart += 2;
                if (_showEnd >= _textObjectStart) _showEnd += 2;
                if (_lastNumberStart >= _textObjectStart) _lastNumberStart += 2;
                if (_ownedLength >= _textObjectStart) _ownedLength += 2;
                _syntheticWeightScope = true;
            }
            TextRenderingMode(2);
            LineWidth(SyntheticBoldStrokeWidth(size));
            LineJoin(1);
            StrokeDash();
            StrokeColor(_paintFillColor);
        } else if (_syntheticBold) {
            TextRenderingMode(0);
        }
        _syntheticBold = bold;
    }
}
