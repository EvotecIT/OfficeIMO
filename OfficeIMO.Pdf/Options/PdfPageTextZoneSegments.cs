namespace OfficeIMO.Pdf;

internal sealed class PdfPageTextZoneSegments {
    internal PdfPageTextZoneSegments(
        System.Collections.Generic.List<PdfFooterSegment>? left,
        System.Collections.Generic.List<PdfFooterSegment>? center,
        System.Collections.Generic.List<PdfFooterSegment>? right) {
        Left = left;
        Center = center;
        Right = right;
    }

    internal System.Collections.Generic.List<PdfFooterSegment>? Left { get; }
    internal System.Collections.Generic.List<PdfFooterSegment>? Center { get; }
    internal System.Collections.Generic.List<PdfFooterSegment>? Right { get; }

    internal bool HasContent =>
        (Left != null && Left.Count > 0) ||
        (Center != null && Center.Count > 0) ||
        (Right != null && Right.Count > 0);

    internal PdfPageTextZoneSegments Clone() => new(
        Left == null ? null : new System.Collections.Generic.List<PdfFooterSegment>(Left),
        Center == null ? null : new System.Collections.Generic.List<PdfFooterSegment>(Center),
        Right == null ? null : new System.Collections.Generic.List<PdfFooterSegment>(Right));
}
