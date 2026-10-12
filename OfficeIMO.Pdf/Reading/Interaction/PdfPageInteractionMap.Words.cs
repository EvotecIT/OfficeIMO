namespace OfficeIMO.Pdf;

public sealed partial class PdfPageInteractionMap {
    /// <summary>Selects the contiguous non-whitespace text run at a visual page coordinate, without crossing a line or a gap between independent runs.</summary>
    public IReadOnlyList<PdfPageInteractionRegion> SelectWord(double x, double y, double tolerance = 0D) {
        // An exact whitespace hit must not be replaced by a neighboring glyph admitted by pointer tolerance.
        PdfPageInteractionRegion? hit = HitTest(x, y).FirstOrDefault(region => region.Kind == PdfInteractionKind.Text)
            ?? HitTest(x, y, tolerance).FirstOrDefault(region => region.Kind == PdfInteractionKind.Text);
        if (hit is null) return Array.Empty<PdfPageInteractionRegion>();
        int index = -1;
        for (int i = 0; i < TextRegions.Count; i++) if (ReferenceEquals(TextRegions[i], hit)) { index = i; break; }
        if (index < 0) return Array.Empty<PdfPageInteractionRegion>();
        if (!IsWordElement(hit)) return Array.Empty<PdfPageInteractionRegion>();
        int first = index, last = index;
        while (first > 0 && IsWordElement(TextRegions[first - 1]) && AreWordNeighbors(TextRegions[first - 1], TextRegions[first])) first--;
        while (last + 1 < TextRegions.Count && IsWordElement(TextRegions[last + 1]) && AreWordNeighbors(TextRegions[last], TextRegions[last + 1])) last++;
        return TextRegions.Skip(first).Take(last - first + 1).ToArray();
    }

    /// <summary>Classifies how <paramref name="next"/> follows <paramref name="previous"/>: touching, after a gap on the same
    /// line, or on another line. Words written as separate runs without space characters, such as an OCR text layer, are
    /// separated by gaps rather than by text.</summary>
    public static PdfTextBreak GetTextBreak(PdfPageInteractionRegion previous, PdfPageInteractionRegion next) {
        Guard.NotNull(previous, nameof(previous));
        Guard.NotNull(next, nameof(next));
        if (AreWordNeighbors(previous, next)) return PdfTextBreak.None;
        return AreOnSameLine(previous.Quad, next.Quad) ? PdfTextBreak.Space : PdfTextBreak.Line;
    }

    private static bool AreOnSameLine(PdfSelectionQuad a, PdfSelectionQuad b) {
        double dx = a.BottomRight.X - a.BottomLeft.X, dy = a.BottomRight.Y - a.BottomLeft.Y;
        double length = Math.Sqrt(dx * dx + dy * dy);
        if (length < 1e-6) { dx = 1D; dy = 0D; length = 1D; }
        // The offset of the next baseline from this one, across the line's direction.
        double offset = ((b.BottomLeft.X - a.BottomLeft.X) * -dy + (b.BottomLeft.Y - a.BottomLeft.Y) * dx) / length;
        double height = Math.Min(Distance(a.TopLeft, a.BottomLeft), Distance(b.TopLeft, b.BottomLeft));
        return Math.Abs(offset) <= Math.Max(0.5D, height * 0.5D);
    }

    private static bool IsWordElement(PdfPageInteractionRegion region) => !string.IsNullOrEmpty(region.Text) && !region.Text!.Any(char.IsWhiteSpace);

    private static bool AreWordNeighbors(PdfPageInteractionRegion first, PdfPageInteractionRegion second) {
        PdfSelectionQuad a = first.Quad, b = second.Quad;
        double height = Math.Min(Distance(a.TopLeft, a.BottomLeft), Distance(b.TopLeft, b.BottomLeft));
        // Below the gap text extraction reads as a word space: about a fifth of the font size, at most two points.
        double tolerance = Math.Max(0.5D, Math.Min(2D, height * 0.15D));
        return (Distance(a.TopRight, b.TopLeft) <= tolerance && Distance(a.BottomRight, b.BottomLeft) <= tolerance) ||
            (Distance(b.TopRight, a.TopLeft) <= tolerance && Distance(b.BottomRight, a.BottomLeft) <= tolerance);
    }

    private static double Distance(PdfSelectionPoint a, PdfSelectionPoint b) => Math.Sqrt((a.X - b.X) * (a.X - b.X) + (a.Y - b.Y) * (a.Y - b.Y));
}
