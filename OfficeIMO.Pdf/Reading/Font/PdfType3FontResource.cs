namespace OfficeIMO.Pdf;

internal sealed class PdfType3FontResource {
    private readonly IReadOnlyDictionary<int, string> _glyphNames;
    private readonly IReadOnlyDictionary<string, PdfStream> _glyphStreams;
    private readonly int _firstCharacter;
    private readonly IReadOnlyList<double> _widths;

    internal PdfType3FontResource(
        Matrix2D fontMatrix,
        PdfDictionary resources,
        bool isUncolored,
        IReadOnlyDictionary<int, string> glyphNames,
        IReadOnlyDictionary<string, PdfStream> glyphStreams,
        int firstCharacter,
        IReadOnlyList<double> widths) {
        FontMatrix = fontMatrix;
        Resources = resources;
        IsUncolored = isUncolored;
        _glyphNames = glyphNames;
        _glyphStreams = glyphStreams;
        _firstCharacter = firstCharacter;
        _widths = widths;
    }

    internal Matrix2D FontMatrix { get; }

    internal PdfDictionary Resources { get; }

    internal bool IsUncolored { get; }

    /// <summary>Estimates the local lookup dictionaries and widths without charging the shared PDF object graph.</summary>
    internal long EstimatedRetainedBytes =>
        (_glyphNames.Count + _glyphStreams.Count) * 96L + _widths.Count * (long)sizeof(double) + 256L;

    internal bool TryGetGlyph(byte characterCode, out PdfStream glyph) {
        string glyphName = _glyphNames.TryGetValue(characterCode, out string? mappedName) ? mappedName : ".notdef";
        if (_glyphStreams.TryGetValue(glyphName, out glyph!)) return true;
        return _glyphStreams.TryGetValue(".notdef", out glyph!);
    }

    internal (double X, double Y) GetGlyphDisplacement(byte characterCode) {
        int index = characterCode - _firstCharacter;
        double width = index >= 0 && index < _widths.Count ? _widths[index] : 0D;
        return (FontMatrix.A * width, FontMatrix.B * width);
    }

    internal double SumNormalizedWidths(byte[]? characterCodes) {
        if (characterCodes == null) return 0D;
        double sum = 0D;
        for (int i = 0; i < characterCodes.Length; i++) {
            int index = characterCodes[i] - _firstCharacter;
            double width = index >= 0 && index < _widths.Count ? _widths[index] : 0D;
            (double x, double y) = GetGlyphDisplacement(characterCodes[i]);
            sum += Math.Sqrt((x * x) + (y * y)) * Math.Sign(width) * 1000D;
        }
        return sum;
    }
}
