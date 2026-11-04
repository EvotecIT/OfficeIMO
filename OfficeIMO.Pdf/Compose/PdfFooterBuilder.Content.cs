namespace OfficeIMO.Pdf;

public sealed partial class PdfFooterBuilder {
    /// <summary>Composes bounded footer content without an additional gap from the body.</summary>
    public PdfFooterBuilder Content(Action<PdfContentBuilder> build, global::OfficeIMO.OfficeLength distanceFromEdge) =>
        Content(build, distanceFromEdge, default);
    /// <summary>Composes bounded footer content without an additional gap from the body.</summary>
    public PdfFooterBuilder Content(Func<PdfRunningContentContext, Action<PdfContentBuilder>> build, global::OfficeIMO.OfficeLength distanceFromEdge) =>
        Content(build, distanceFromEdge, default);
    /// <summary>Composes bounded footer content without an additional gap from the body.</summary>
    public PdfFooterBuilder FirstPageContent(Action<PdfContentBuilder> build, global::OfficeIMO.OfficeLength distanceFromEdge) =>
        FirstPageContent(build, distanceFromEdge, default);
    /// <summary>Composes bounded footer content without an additional gap from the body.</summary>
    public PdfFooterBuilder FirstPageContent(Func<PdfRunningContentContext, Action<PdfContentBuilder>> build, global::OfficeIMO.OfficeLength distanceFromEdge) =>
        FirstPageContent(build, distanceFromEdge, default);
    /// <summary>Composes bounded footer content without an additional gap from the body.</summary>
    public PdfFooterBuilder EvenPagesContent(Action<PdfContentBuilder> build, global::OfficeIMO.OfficeLength distanceFromEdge) =>
        EvenPagesContent(build, distanceFromEdge, default);
    /// <summary>Composes bounded footer content without an additional gap from the body.</summary>
    public PdfFooterBuilder EvenPagesContent(Func<PdfRunningContentContext, Action<PdfContentBuilder>> build, global::OfficeIMO.OfficeLength distanceFromEdge) =>
        EvenPagesContent(build, distanceFromEdge, default);

    /// <summary>Replaces the default footer with bounded document flow at a distance from the bottom page edge. Uses normal content-builder styles; content cannot create pages.</summary>
    public PdfFooterBuilder Content(Action<PdfContentBuilder> build, global::OfficeIMO.OfficeLength distanceFromEdge, global::OfficeIMO.OfficeLength bodyGap) =>
        SetContent(build, distanceFromEdge.ToPoints(), bodyGap.ToPoints(), PdfRunningContentVariant.Default);
    /// <summary>Composes the default bounded footer from final page information. The factory can run during pagination stabilization and must be deterministic.</summary>
    public PdfFooterBuilder Content(Func<PdfRunningContentContext, Action<PdfContentBuilder>> build, global::OfficeIMO.OfficeLength distanceFromEdge, global::OfficeIMO.OfficeLength bodyGap) =>
        SetContent(build, distanceFromEdge.ToPoints(), bodyGap.ToPoints(), PdfRunningContentVariant.Default);
    /// <summary>Replaces the first-page footer with bounded flow measured from the bottom page edge.</summary>
    public PdfFooterBuilder FirstPageContent(Action<PdfContentBuilder> build, global::OfficeIMO.OfficeLength distanceFromEdge, global::OfficeIMO.OfficeLength bodyGap) =>
        SetContent(build, distanceFromEdge.ToPoints(), bodyGap.ToPoints(), PdfRunningContentVariant.FirstPage);
    /// <summary>Composes bounded first-page footer flow from final page information.</summary>
    public PdfFooterBuilder FirstPageContent(Func<PdfRunningContentContext, Action<PdfContentBuilder>> build, global::OfficeIMO.OfficeLength distanceFromEdge, global::OfficeIMO.OfficeLength bodyGap) =>
        SetContent(build, distanceFromEdge.ToPoints(), bodyGap.ToPoints(), PdfRunningContentVariant.FirstPage);
    /// <summary>Replaces even-page footers with bounded flow measured from the bottom page edge.</summary>
    public PdfFooterBuilder EvenPagesContent(Action<PdfContentBuilder> build, global::OfficeIMO.OfficeLength distanceFromEdge, global::OfficeIMO.OfficeLength bodyGap) =>
        SetContent(build, distanceFromEdge.ToPoints(), bodyGap.ToPoints(), PdfRunningContentVariant.EvenPages);
    /// <summary>Composes bounded even-page footer flow from final page information.</summary>
    public PdfFooterBuilder EvenPagesContent(Func<PdfRunningContentContext, Action<PdfContentBuilder>> build, global::OfficeIMO.OfficeLength distanceFromEdge, global::OfficeIMO.OfficeLength bodyGap) =>
        SetContent(build, distanceFromEdge.ToPoints(), bodyGap.ToPoints(), PdfRunningContentVariant.EvenPages);

    private PdfFooterBuilder SetContent(Action<PdfContentBuilder> build, double distance, double gap, PdfRunningContentVariant variant) {
        Guard.NotNull(build, nameof(build));
        Guard.NonNegative(distance, nameof(distance)); Guard.NonNegative(gap, nameof(gap));
        IReadOnlyList<IPdfBlock> blocks = _doc.BuildFlowBlocks(build);
        _opts.SetRunningContent(new PdfRunningContent(_ => blocks, distance, gap, usesPageContext: false), false, variant);
        return this;
    }

    private PdfFooterBuilder SetContent(Func<PdfRunningContentContext, Action<PdfContentBuilder>> build, double distance, double gap, PdfRunningContentVariant variant) {
        Guard.NotNull(build, nameof(build));
        Guard.NonNegative(distance, nameof(distance)); Guard.NonNegative(gap, nameof(gap));
        _opts.SetRunningContent(new PdfRunningContent(context => _doc.BuildFlowBlocks(build(context)
            ?? throw new InvalidOperationException("Running PDF content factory returned null.")), distance, gap, usesPageContext: true), false, variant);
        return this;
    }
}
