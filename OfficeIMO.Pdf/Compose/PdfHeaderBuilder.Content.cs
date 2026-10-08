namespace OfficeIMO.Pdf;

public sealed partial class PdfHeaderBuilder {
    /// <summary>Replaces the default header with bounded document flow at a distance from the top page edge. Uses normal content-builder styles; content cannot create pages.</summary>
    public PdfHeaderBuilder Content(Action<PdfContentBuilder> build, double distanceFromEdge, double bodyGap = 0D) =>
        SetContent(build, distanceFromEdge, bodyGap, PdfRunningContentVariant.Default);
    /// <summary>Composes the default bounded header from final page information. The factory can run during pagination stabilization and must be deterministic.</summary>
    public PdfHeaderBuilder Content(Func<PdfRunningContentContext, Action<PdfContentBuilder>> build, double distanceFromEdge, double bodyGap = 0D) =>
        SetContent(build, distanceFromEdge, bodyGap, PdfRunningContentVariant.Default);
    /// <summary>Replaces the first-page header with bounded flow measured from the top page edge.</summary>
    public PdfHeaderBuilder FirstPageContent(Action<PdfContentBuilder> build, double distanceFromEdge, double bodyGap = 0D) =>
        SetContent(build, distanceFromEdge, bodyGap, PdfRunningContentVariant.FirstPage);
    /// <summary>Composes bounded first-page header flow from final page information.</summary>
    public PdfHeaderBuilder FirstPageContent(Func<PdfRunningContentContext, Action<PdfContentBuilder>> build, double distanceFromEdge, double bodyGap = 0D) =>
        SetContent(build, distanceFromEdge, bodyGap, PdfRunningContentVariant.FirstPage);
    /// <summary>Replaces even-page headers with bounded flow measured from the top page edge.</summary>
    public PdfHeaderBuilder EvenPagesContent(Action<PdfContentBuilder> build, double distanceFromEdge, double bodyGap = 0D) =>
        SetContent(build, distanceFromEdge, bodyGap, PdfRunningContentVariant.EvenPages);
    /// <summary>Composes bounded even-page header flow from final page information.</summary>
    public PdfHeaderBuilder EvenPagesContent(Func<PdfRunningContentContext, Action<PdfContentBuilder>> build, double distanceFromEdge, double bodyGap = 0D) =>
        SetContent(build, distanceFromEdge, bodyGap, PdfRunningContentVariant.EvenPages);

    private PdfHeaderBuilder SetContent(Action<PdfContentBuilder> build, double distance, double gap, PdfRunningContentVariant variant) {
        Guard.NotNull(build, nameof(build));
        Guard.NonNegative(distance, nameof(distance)); Guard.NonNegative(gap, nameof(gap));
        IReadOnlyList<IPdfBlock> blocks = _doc.BuildFlowBlocks(build);
        _opts.SetRunningContent(new PdfRunningContent(_ => blocks, distance, gap, usesPageContext: false), true, variant);
        return this;
    }

    private PdfHeaderBuilder SetContent(Func<PdfRunningContentContext, Action<PdfContentBuilder>> build, double distance, double gap, PdfRunningContentVariant variant) {
        Guard.NotNull(build, nameof(build));
        _opts.SetRunningContent(new PdfRunningContent(context => _doc.BuildFlowBlocks(build(context)
            ?? throw new InvalidOperationException("Running PDF content factory returned null.")), distance, gap, usesPageContext: true), true, variant);
        return this;
    }
}
