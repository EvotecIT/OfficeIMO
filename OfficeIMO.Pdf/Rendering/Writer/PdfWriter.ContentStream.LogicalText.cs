namespace OfficeIMO.Pdf;

internal sealed partial class ContentStreamBuilder {
    private double _textA = 1, _textB, _textC, _textD = 1, _textE, _textF, _lineE, _lineF;
    private double _textScale = 1, _textLeading, _textWordSpacing;
    private const double SyntheticObliqueShear = 1D / 3D;
    private bool _syntheticOblique, _hasTextMatrix;
    private bool _isolatedText;
    private readonly Stack<(double Scale, double Leading, double WordSpacing, bool SyntheticOblique)> _textStates = new();

    private void ResetTrackedTextMatrix() {
        _textA = _textD = 1; _textB = _textC = _textE = _textF = _lineE = _lineF = 0;
        _isolatedText = false;
        _hasTextMatrix = false;
    }

    private void TrackTextMatrix(double a, double b, double c, double d, double e, double f) {
        _hasTextMatrix = true;
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
            PdfGlyphInfo single = glyphs[index];
            if (!single.HasPositioning &&
                (index + 1 == glyphs.Count || glyphs[index + 1].LogicalClusterStart != single.LogicalClusterStart)) {
                // Keep the same isolated text object and ActualText boundary, but
                // avoid allocating a list, builders and formatted strings per scalar.
                _sb.Append("ET\nBT\n");
                TextMatrixApplied(_textA, _textB, _textC, _textD, _textE, _textF);
                bool singleMarked = single.UnicodeText.Length != 0;
                if (singleMarked) {
                    _sb.Append("/Span << /ActualText ");
                    PdfSyntaxEscaper.AppendTextStringCancellable(_sb, single.UnicodeText, default);
                    _sb.Append(" >> BDC\n");
                }
                _sb.Append('<');
                PdfGlyphRun.AppendGlyphHex(_sb, single.GlyphId);
                _sb.Append("> Tj\n");
                if (singleMarked) _sb.Append("EMC\n");
                AdvanceTrackedText(single.AdvanceWidth1000 * fontSize / 1000D);
                index++;
                continue;
            }
            var cluster = new List<PdfGlyphInfo>();
            var logical = new System.Text.StringBuilder();
            int logicalEnd = glyphs[index].LogicalClusterStart;
            do {
                PdfGlyphInfo glyph = glyphs[index++];
                cluster.Add(glyph);
                logical.Append(glyph.UnicodeText);
                // GSUB multiple substitution and a later ligature can give
                // neighboring glyphs different starts but overlapping source
                // ownership. Keep that whole interval in one redaction scope.
                logicalEnd = Math.Max(logicalEnd, Math.Max(glyph.LogicalClusterStart + 1,
                    glyph.TextIndex + glyph.UnicodeText.Length));
            } while (index < glyphs.Count && glyphs[index].LogicalClusterStart < logicalEnd);
            _sb.Append("ET\nBT\n");
            TextMatrixApplied(_textA, _textB, _textC, _textD, _textE, _textF);
            bool marked = logical.Length != 0;
            if (marked) _sb.Append("/Span << /ActualText ").Append(PdfSyntaxEscaper.TextString(logical.ToString())).Append(" >> BDC\n");
            if (cluster.Any(glyph => glyph.HasPositioning)) AppendPositionedGlyphs(cluster, fontSize, textRise);
            else ShowHexText(string.Concat(cluster.Select(glyph => glyph.GlyphId.ToString("X4", System.Globalization.CultureInfo.InvariantCulture))));
            if (marked) _sb.Append("EMC\n");
            AdvanceTrackedText(cluster.Sum(glyph => glyph.AdvanceWidth1000) * fontSize / 1000D);
        }
        _sb.Append("ET\nBT\n");
        TextMatrixApplied(_textA, _textB, _textC, _textD, _textE, _textF);
        _lineE = lineE; _lineF = lineF; _isolatedText = true;
    }

}
