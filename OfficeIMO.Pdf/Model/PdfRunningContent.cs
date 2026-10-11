namespace OfficeIMO.Pdf;

internal enum PdfRunningContentVariant { Default, FirstPage, EvenPages }

// Immutable definition. Materializations and layout resources belong to one serialization,
// rather than to the document options or a user-supplied factory's lifetime.
internal sealed class PdfRunningContent {
    internal PdfRunningContent(Func<PdfRunningContentContext, IReadOnlyList<IPdfBlock>> materialize,
        double distanceFromEdge, double bodyGap, bool usesPageContext, bool allowTextFallback = false,
        bool reserveDistanceWhenEmpty = false, bool retainTextProjection = false) {
        Guard.NotNull(materialize, nameof(materialize));
        Guard.NonNegative(distanceFromEdge, nameof(distanceFromEdge));
        Guard.NonNegative(bodyGap, nameof(bodyGap));
        Materialize = materialize;
        DistanceFromEdge = distanceFromEdge;
        BodyGap = bodyGap;
        UsesPageContext = usesPageContext;
        AllowTextFallback = allowTextFallback;
        ReserveDistanceWhenEmpty = reserveDistanceWhenEmpty;
        RetainTextProjection = retainTextProjection;
    }

    internal Func<PdfRunningContentContext, IReadOnlyList<IPdfBlock>> Materialize { get; }
    internal double DistanceFromEdge { get; }
    internal double BodyGap { get; }
    internal bool UsesPageContext { get; }
    // Imported ordinary stories also retain a styled text projection. Decide
    // between them with this page's final fields and measured physical frame.
    internal bool AllowTextFallback { get; }
    // A referenced imported story still owns its edge distance when all of its
    // paragraph marks are hidden. Missing stories keep the ordinary empty-flow contract.
    internal bool ReserveDistanceWhenEmpty { get; }
    // A geometry-only imported story can reserve its source line while the
    // caller-configured text projection retains its physical-margin anchor.
    internal bool RetainTextProjection { get; }
    internal PdfRunningContent WithTextFallback() =>
        new(Materialize, DistanceFromEdge, BodyGap, UsesPageContext, allowTextFallback: true,
            reserveDistanceWhenEmpty: ReserveDistanceWhenEmpty, retainTextProjection: RetainTextProjection);
    internal PdfRunningContent WithReservedEmptyDistance() =>
        new(Materialize, DistanceFromEdge, BodyGap, UsesPageContext, AllowTextFallback,
            reserveDistanceWhenEmpty: true, retainTextProjection: RetainTextProjection);
    internal PdfRunningContent WithTextProjection() =>
        new(Materialize, DistanceFromEdge, BodyGap, UsesPageContext, AllowTextFallback,
            ReserveDistanceWhenEmpty, retainTextProjection: true);
}
