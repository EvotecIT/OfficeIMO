namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private bool _fontLayoutRebuilt;

    /// <summary>
    /// Selects the approved default face from laid-out flow before page admission
    /// and print fitting. Rebuild once with the frozen face, reusing the resource
    /// session, diagnostics and cumulative operation budget. The new engine owns
    /// its own pagination convergence and measurement caches.
    /// </summary>
    private bool TryRebuildWithPreparedFontMetrics(IReadOnlyList<HtmlRenderFlowBlock> blocks,
        out HtmlRenderDocument rendered) {
        rendered = null!;
        if (_options is not Pdf.HtmlToPdfOptions pdfOptions
            || !Pdf.HtmlPdfRenderedConverter.TryPrepareLibraryFontMetrics(
                blocks.SelectMany(block => block.Visuals), _fonts, pdfOptions, _cancellationToken)) return false;
        CheckCancellation();
        _fontLayoutRebuilt = true;
        rendered = new HtmlRenderLayoutEngine(_document, _computedStyles, _options,
            _diagnostics, _resources, _pageRules, _fonts, _limits,
            _initialLogicalTextOrder, _initialSemanticNodeId, _operationBudget, _cancellationToken, _fontUsage).Render();
        return true;
    }
}
