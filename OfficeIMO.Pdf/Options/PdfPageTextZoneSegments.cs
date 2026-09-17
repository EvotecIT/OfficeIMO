namespace OfficeIMO.Pdf;

internal sealed class PdfPageTextZoneSegments {
    internal PdfPageTextZoneSegments(
        System.Collections.Generic.List<FooterSegment>? left,
        System.Collections.Generic.List<FooterSegment>? center,
        System.Collections.Generic.List<FooterSegment>? right) {
        Left = left;
        Center = center;
        Right = right;
    }

    internal System.Collections.Generic.List<FooterSegment>? Left { get; }
    internal System.Collections.Generic.List<FooterSegment>? Center { get; }
    internal System.Collections.Generic.List<FooterSegment>? Right { get; }

    internal bool HasContent =>
        (Left != null && Left.Count > 0) ||
        (Center != null && Center.Count > 0) ||
        (Right != null && Right.Count > 0);

    internal PdfPageTextZoneSegments Clone() => new(
        Left == null ? null : new System.Collections.Generic.List<FooterSegment>(Left),
        Center == null ? null : new System.Collections.Generic.List<FooterSegment>(Center),
        Right == null ? null : new System.Collections.Generic.List<FooterSegment>(Right));
}
