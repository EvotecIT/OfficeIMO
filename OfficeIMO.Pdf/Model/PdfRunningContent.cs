namespace OfficeIMO.Pdf;

internal enum PdfRunningContentVariant { Default, FirstPage, EvenPages }

// Immutable definition. Materializations and layout resources belong to one serialization,
// rather than to the document options or a user-supplied factory's lifetime.
internal sealed class PdfRunningContent {
    internal PdfRunningContent(Func<PdfRunningContentContext, IReadOnlyList<IPdfBlock>> materialize,
        double distanceFromEdge, double bodyGap, bool usesPageContext) {
        Guard.NotNull(materialize, nameof(materialize));
        Guard.NonNegative(distanceFromEdge, nameof(distanceFromEdge));
        Guard.NonNegative(bodyGap, nameof(bodyGap));
        Materialize = materialize;
        DistanceFromEdge = distanceFromEdge;
        BodyGap = bodyGap;
        UsesPageContext = usesPageContext;
    }

    internal Func<PdfRunningContentContext, IReadOnlyList<IPdfBlock>> Materialize { get; }
    internal double DistanceFromEdge { get; }
    internal double BodyGap { get; }
    internal bool UsesPageContext { get; }
}
