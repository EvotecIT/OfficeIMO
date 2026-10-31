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

    internal PdfRedactionGlyphSelection(IEnumerable<PdfRedactionSelectedGlyph> glyphs, bool isHiddenOnly = false, bool isCompleteBlock = false) {
        _glyphs = glyphs.ToArray();
        IsHiddenOnly = isHiddenOnly;
        IsCompleteBlock = isCompleteBlock;
    }

    /// <summary>True when the selection is a whole logical block rather than a matched range.</summary>
    internal bool IsCompleteBlock { get; }

    internal int Count => _glyphs.Length;

    /// <summary>True when nothing visible is selected, because every glyph is in a hidden layer or there are none, so no mark or underlay removal is needed.</summary>
    internal bool IsHiddenOnly { get; }

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

    /// <summary>True when the rectangle overlaps the ink of a selected glyph, as opposed to the mark around it.</summary>
    internal bool OverlapsInk(double x, double y, double width, double height) {
        for (int index = 0; index < _glyphs.Length; index++) {
            PdfTextSpanBounds ink = _glyphs[index].Ink;
            if (x < ink.Right && x + width > ink.Left && y < ink.Top && y + height > ink.Bottom) return true;
        }
        return false;
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

    internal PdfRedactionSelectedGlyph(double x, double y, IReadOnlyList<byte> encodedBytes, PdfTextSpanBounds ink) {
        X = x;
        Y = y;
        Ink = ink;
        _encodedBytes = encodedBytes.ToArray();
    }

    internal PdfTextSpanBounds Ink { get; }

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
