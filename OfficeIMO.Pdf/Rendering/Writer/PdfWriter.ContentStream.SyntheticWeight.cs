namespace OfficeIMO.Pdf;

internal sealed partial class ContentStreamBuilder {
    private OfficeColor _paintFillColor = OfficeColor.Black;

    internal static double SyntheticBoldStrokeWidth(double size) => size / 24D;

    private void ApplySyntheticWeight(double size, bool bold) {
        if (bold) {
            if (!T._inText) throw new InvalidOperationException("Synthetic font weight requires a text object.");
            if (!T._syntheticWeightScope) {
                // The need for synthesis may first appear halfway through a rich
                // paragraph. Scope its entire text object, including already written
                // runs, so stroke width, join and colour cannot leak into later paint.
                _sb.Insert(T._textObjectStart, "q\n");
                if (T._showStart >= T._textObjectStart) T._showStart += 2;
                if (T._showEnd >= T._textObjectStart) T._showEnd += 2;
                if (T._lastNumberStart >= T._textObjectStart) T._lastNumberStart += 2;
                if (T._ownedLength >= T._textObjectStart) T._ownedLength += 2;
                T._syntheticWeightScope = true;
            }
            TextRenderingMode(2);
            LineWidth(SyntheticBoldStrokeWidth(size));
            LineJoin(1);
            StrokeDash();
            StrokeColor(_paintFillColor);
        } else if (T._syntheticBold) {
            TextRenderingMode(0);
        }
        T._syntheticBold = bold;
    }
}
