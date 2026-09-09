namespace OfficeIMO.Pdf;

public sealed partial class PdfPageInteractionMap {
    /// <summary>Selects the contiguous non-whitespace text run at a visual page coordinate, without crossing a line or a gap between independent runs.</summary>
    public IReadOnlyList<PdfPageInteractionRegion> SelectWord(double x, double y, double tolerance = 0D) {
        PdfPageInteractionRegion? hit = HitTest(x, y, tolerance).FirstOrDefault(region => region.Kind == PdfInteractionKind.Text);
        if (hit is null) return Array.Empty<PdfPageInteractionRegion>();
        int index = -1;
        for (int i = 0; i < TextRegions.Count; i++) if (ReferenceEquals(TextRegions[i], hit)) { index = i; break; }
        if (index < 0) return Array.Empty<PdfPageInteractionRegion>();
        if (!IsWordElement(hit)) return new[] { hit };
        int first = index, last = index;
        while (first > 0 && IsWordElement(TextRegions[first - 1]) && AreWordNeighbors(TextRegions[first - 1], TextRegions[first])) first--;
        while (last + 1 < TextRegions.Count && IsWordElement(TextRegions[last + 1]) && AreWordNeighbors(TextRegions[last], TextRegions[last + 1])) last++;
        return TextRegions.Skip(first).Take(last - first + 1).ToArray();
    }

    private static bool IsWordElement(PdfPageInteractionRegion region) => !string.IsNullOrEmpty(region.Text) && !region.Text!.Any(char.IsWhiteSpace);

    private static bool AreWordNeighbors(PdfPageInteractionRegion first, PdfPageInteractionRegion second) {
        PdfSelectionQuad a = first.Quad, b = second.Quad;
        double height = Math.Min(Distance(a.TopLeft, a.BottomLeft), Distance(b.TopLeft, b.BottomLeft));
        double tolerance = Math.Max(0.5D, height * 0.15D);
        return (Distance(a.TopRight, b.TopLeft) <= tolerance && Distance(a.BottomRight, b.BottomLeft) <= tolerance) ||
            (Distance(b.TopRight, a.TopLeft) <= tolerance && Distance(b.BottomRight, a.BottomLeft) <= tolerance);
    }

    private static double Distance(PdfSelectionPoint a, PdfSelectionPoint b) => Math.Sqrt((a.X - b.X) * (a.X - b.X) + (a.Y - b.Y) * (a.Y - b.Y));
}
