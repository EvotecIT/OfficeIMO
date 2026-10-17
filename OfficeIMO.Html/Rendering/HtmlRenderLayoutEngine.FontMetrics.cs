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
                blocks.SelectMany(block => block.Visuals), EnumerateMarginFontRequirements(blocks),
                _fonts, pdfOptions, _cancellationToken)) return false;
        CheckCancellation();
        _fontLayoutRebuilt = true;
        rendered = new HtmlRenderLayoutEngine(_document, _computedStyles, _options,
            _diagnostics, _resources, _pageRules, _fonts, _limits,
            _initialLogicalTextOrder, _initialSemanticNodeId, _operationBudget, _cancellationToken, _fontUsage).Render();
        return true;
    }

    /// <summary>
    /// Includes declared repeated page text before pagination. First, left and
    /// later right pages cover the supported page selectors; counters need only
    /// representative values for font coverage. Actual margin layout still owns
    /// page counts, running strings, clipping and diagnostics.
    /// </summary>
    private IEnumerable<(string Text, OfficeFontInfo Font)> EnumerateMarginFontRequirements(
        IReadOnlyList<HtmlRenderFlowBlock> blocks) {
        if (_options.Mode != HtmlRenderMode.Paged) yield break;
        foreach (string? pageName in blocks.Select(block => block.PageName)
                     .Concat(blocks.SelectMany(block => block.ForcedBreaks.Select(item => item.PageName)))
                     .Concat(new string?[] { null }).Distinct(StringComparer.Ordinal)) {
            for (int pageNumber = 1; pageNumber <= 3; pageNumber++) {
                CheckCancellation();
                ChargeLayoutOperations(2L * _pageRules.RuleCount, "page-margin font rule matching");
                HtmlCssPageGeometry geometry = _pageRules.ResolveGeometry(pageNumber, pageName, _options);
                foreach (HtmlCssPageMarginTemplate box in _pageRules.ResolveMarginBoxes(
                             pageNumber, pageName, geometry, _options).Values) {
                    if (box.Content.ContainsRunningElement) continue;
                    ChargeLayoutOperations(box.Content.GetRenderedLength(pageNumber, 3, null), "page-margin font requirements");
                    string text = box.Content.Render(pageNumber, 3, null);
                    if (text.Length > 0) yield return (text, box.Font);
                }
            }
        }
    }
}
