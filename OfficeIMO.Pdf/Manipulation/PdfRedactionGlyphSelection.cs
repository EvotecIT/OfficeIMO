namespace OfficeIMO.Pdf;

/// <summary>
/// The exact encoded glyphs a text search selected for one redaction area. A search mark keeps the full
/// glyph envelope so it hides ascenders and descenders, and that envelope can reach into a neighbouring
/// line. Removal therefore follows this selection, never the rectangle.
/// </summary>
/// <remarks>
/// A glyph is identified by its page-space origin and encoded bytes. Both survive the content offset
/// changes made by earlier rewrite steps, and the search planner refuses any unselected text that
/// coincides with a selected glyph, so the pair cannot name a second glyph on that page.
/// </remarks>
internal sealed class PdfRedactionGlyphSelection {
    private const double OriginTolerance = 0.005D;
    private readonly PdfRedactionSelectedGlyph[] _glyphs;

    internal PdfRedactionGlyphSelection(IEnumerable<PdfRedactionSelectedGlyph> glyphs, IEnumerable<string> matchedTexts) {
        _glyphs = glyphs.ToArray();
        MatchedTexts = matchedTexts.Distinct(StringComparer.Ordinal).ToArray();
    }

    internal int Count => _glyphs.Length;

    /// <summary>The matched wording, used to clear copies of it held outside the page content.</summary>
    internal IReadOnlyList<string> MatchedTexts { get; }

    internal bool Contains(double x, double y, IReadOnlyList<byte> encodedBytes) {
        for (int index = 0; index < _glyphs.Length; index++) {
            if (_glyphs[index].Matches(x, y, encodedBytes, OriginTolerance)) return true;
        }
        return false;
    }

    /// <summary>Records every selected glyph resolved by one glyph found in a document.</summary>
    internal void MarkResolved(double x, double y, IReadOnlyList<byte> encodedBytes, ISet<int> resolved) {
        for (int index = 0; index < _glyphs.Length; index++) {
            if (_glyphs[index].Matches(x, y, encodedBytes, OriginTolerance)) resolved.Add(index);
        }
    }

    internal bool HasOriginWithin(PdfTextSpanBounds bounds) {
        for (int index = 0; index < _glyphs.Length; index++) {
            PdfRedactionSelectedGlyph glyph = _glyphs[index];
            if (glyph.X >= bounds.Left - OriginTolerance && glyph.X <= bounds.Right + OriginTolerance &&
                glyph.Y >= bounds.Bottom - OriginTolerance && glyph.Y <= bounds.Top + OriginTolerance) return true;
        }
        return false;
    }
}

internal readonly struct PdfRedactionSelectedGlyph {
    private readonly byte[] _encodedBytes;

    internal PdfRedactionSelectedGlyph(double x, double y, IReadOnlyList<byte> encodedBytes) {
        X = x;
        Y = y;
        _encodedBytes = encodedBytes.ToArray();
    }

    internal double X { get; }
    internal double Y { get; }

    internal bool Matches(double x, double y, IReadOnlyList<byte> encodedBytes, double tolerance) {
        if (Math.Abs(X - x) > tolerance || Math.Abs(Y - y) > tolerance || _encodedBytes.Length != encodedBytes.Count) return false;
        for (int index = 0; index < _encodedBytes.Length; index++) {
            if (_encodedBytes[index] != encodedBytes[index]) return false;
        }
        return true;
    }
}
