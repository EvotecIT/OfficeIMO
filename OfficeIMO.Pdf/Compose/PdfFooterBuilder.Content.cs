namespace OfficeIMO.Pdf;

public sealed partial class PdfFooterBuilder {
    /// <summary>Replaces the default footer with bounded document flow at a distance from the bottom page edge. Uses normal content-builder styles; content cannot create pages.</summary>
    public PdfFooterBuilder Content(Action<PdfContentBuilder> build, double distanceFromEdge, double bodyGap = 0D) =>
        SetContent(build, distanceFromEdge, bodyGap, PdfRunningContentVariant.Default);
    /// <summary>Composes the default bounded footer from final page information. The factory can run during pagination stabilization and must be deterministic.</summary>
    public PdfFooterBuilder Content(Func<PdfRunningContentContext, Action<PdfContentBuilder>> build, double distanceFromEdge, double bodyGap = 0D) =>
        SetContent(build, distanceFromEdge, bodyGap, PdfRunningContentVariant.Default);
    /// <summary>Replaces the first-page footer with bounded flow measured from the bottom page edge.</summary>
    public PdfFooterBuilder FirstPageContent(Action<PdfContentBuilder> build, double distanceFromEdge, double bodyGap = 0D) =>
        SetContent(build, distanceFromEdge, bodyGap, PdfRunningContentVariant.FirstPage);
    /// <summary>Composes bounded first-page footer flow from final page information.</summary>
    public PdfFooterBuilder FirstPageContent(Func<PdfRunningContentContext, Action<PdfContentBuilder>> build, double distanceFromEdge, double bodyGap = 0D) =>
        SetContent(build, distanceFromEdge, bodyGap, PdfRunningContentVariant.FirstPage);
    /// <summary>Replaces even-page footers with bounded flow measured from the bottom page edge.</summary>
    public PdfFooterBuilder EvenPagesContent(Action<PdfContentBuilder> build, double distanceFromEdge, double bodyGap = 0D) =>
        SetContent(build, distanceFromEdge, bodyGap, PdfRunningContentVariant.EvenPages);
    /// <summary>Composes bounded even-page footer flow from final page information.</summary>
    public PdfFooterBuilder EvenPagesContent(Func<PdfRunningContentContext, Action<PdfContentBuilder>> build, double distanceFromEdge, double bodyGap = 0D) =>
        SetContent(build, distanceFromEdge, bodyGap, PdfRunningContentVariant.EvenPages);

    private PdfFooterBuilder SetContent(Action<PdfContentBuilder> build, double distance, double gap, PdfRunningContentVariant variant) {
        Guard.NotNull(build, nameof(build));
        Guard.NonNegative(distance, nameof(distance)); Guard.NonNegative(gap, nameof(gap));
        IReadOnlyList<IPdfBlock> blocks = _doc.BuildFlowBlocks(build);
        _opts.SetRunningContent(new PdfRunningContent(_ => blocks, distance, gap, usesPageContext: false), false, variant);
        return this;
    }

    private PdfFooterBuilder SetContent(Func<PdfRunningContentContext, Action<PdfContentBuilder>> build, double distance, double gap, PdfRunningContentVariant variant) {
        Guard.NotNull(build, nameof(build));
        _opts.SetRunningContent(new PdfRunningContent(context => _doc.BuildFlowBlocks(build(context)
            ?? throw new InvalidOperationException("Running PDF content factory returned null.")), distance, gap, usesPageContext: true), false, variant);
        return this;
    }
}
