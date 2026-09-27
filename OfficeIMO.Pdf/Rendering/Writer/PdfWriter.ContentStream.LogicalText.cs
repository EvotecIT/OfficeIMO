namespace OfficeIMO.Pdf;

internal sealed partial class ContentStreamBuilder {
    private double _textA = 1, _textB, _textC, _textD = 1, _textE, _textF, _lineE, _lineF;
    private double _textScale = 1, _textLeading, _textWordSpacing;
    private bool _isolatedText;
    private readonly Stack<(double Scale, double Leading, double WordSpacing)> _textStates = new();

    private void ResetTrackedTextMatrix() {
        _textA = _textD = 1; _textB = _textC = _textE = _textF = _lineE = _lineF = 0;
        _isolatedText = false;
    }

    private void TrackTextMatrix(double a, double b, double c, double d, double e, double f) {
        _textA = a; _textB = b; _textC = c; _textD = d;
        _textE = _lineE = e; _textF = _lineF = f;
    }

    private void AdvanceTrackedText(double advance) {
        _textE += _textA * advance * _textScale;
        _textF += _textB * advance * _textScale;
    }

    // Isolate source clusters so conservative ActualText redaction cannot discard
    // independent neighboring source characters.
    private void WriteIsolatedLogicalGlyphs(IReadOnlyList<PdfGlyphInfo> glyphs, double fontSize, double textRise) {
        double lineE = _lineE, lineF = _lineF;
        for (int index = 0; index < glyphs.Count;) {
            var cluster = new List<PdfGlyphInfo>();
            var logical = new System.Text.StringBuilder();
            int clusterStart = glyphs[index].LogicalClusterStart;
            do {
                PdfGlyphInfo glyph = glyphs[index++];
                cluster.Add(glyph);
                logical.Append(glyph.UnicodeText);
            } while (index < glyphs.Count && glyphs[index].LogicalClusterStart == clusterStart);
            _sb.Append("ET\nBT\n");
            TextMatrix(_textA, _textB, _textC, _textD, _textE, _textF);
            bool marked = logical.Length != 0;
            if (marked) _sb.Append("/Span << /ActualText ").Append(PdfSyntaxEscaper.TextString(logical.ToString())).Append(" >> BDC\n");
            if (cluster.Any(glyph => glyph.HasPositioning)) AppendPositionedGlyphs(cluster, fontSize, textRise);
            else ShowHexText(string.Concat(cluster.Select(glyph => glyph.GlyphId.ToString("X4", System.Globalization.CultureInfo.InvariantCulture))));
            if (marked) _sb.Append("EMC\n");
            AdvanceTrackedText(cluster.Sum(glyph => glyph.AdvanceWidth1000) * fontSize / 1000D);
        }
        _sb.Append("ET\nBT\n");
        TextMatrix(_textA, _textB, _textC, _textD, _textE, _textF);
        _lineE = lineE; _lineF = lineF; _isolatedText = true;
    }

}
