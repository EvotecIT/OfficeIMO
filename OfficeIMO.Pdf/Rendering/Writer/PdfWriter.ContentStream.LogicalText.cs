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

    // Isolate logical words so glyph aliases retain their exact source text and
    // conservative ActualText redaction cannot discard neighboring words.
    private void WriteIsolatedLogicalGlyphs(IReadOnlyList<PdfGlyphInfo> glyphs, double fontSize, double textRise) {
        double lineE = _lineE, lineF = _lineF;
        for (int index = 0; index < glyphs.Count;) {
            var word = new List<PdfGlyphInfo>();
            var logical = new System.Text.StringBuilder();
            bool whitespace = string.IsNullOrWhiteSpace(glyphs[index].UnicodeText);
            do {
                PdfGlyphInfo glyph = glyphs[index++];
                word.Add(glyph);
                logical.Append(glyph.UnicodeText);
            } while (index < glyphs.Count && string.IsNullOrWhiteSpace(glyphs[index].UnicodeText) == whitespace);
            _sb.Append("ET\nBT\n");
            TextMatrix(_textA, _textB, _textC, _textD, _textE, _textF);
            bool marked = logical.Length != 0;
            if (marked) _sb.Append("/Span << /ActualText ").Append(PdfSyntaxEscaper.TextString(logical.ToString())).Append(" >> BDC\n");
            if (word.Any(glyph => glyph.HasPositioning)) AppendPositionedGlyphs(word, fontSize, textRise);
            else ShowHexText(string.Concat(word.Select(glyph => glyph.GlyphId.ToString("X4", System.Globalization.CultureInfo.InvariantCulture))));
            if (marked) _sb.Append("EMC\n");
            AdvanceTrackedText(word.Sum(glyph => glyph.AdvanceWidth1000) * fontSize / 1000D);
        }
        _sb.Append("ET\nBT\n");
        TextMatrix(_textA, _textB, _textC, _textD, _textE, _textF);
        _lineE = lineE; _lineF = lineF; _isolatedText = true;
    }
}
