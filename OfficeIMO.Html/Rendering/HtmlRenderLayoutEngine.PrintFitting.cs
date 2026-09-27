namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private int _printFitPasses;

    private bool CapturePrintLayoutBoxes => _options.AutoFitWidePrintRoot
        && _options.Mode == HtmlRenderMode.Paged && _options.HonorCssPageRules
        && !_pageRules.HasPageSpecificRules;

    private HtmlRenderDocument CompletePrintLayout(HtmlRenderDocument rendered) {
        if (!CapturePrintLayoutBoxes) return rendered;
        double contentWidth = _options.PageWidth - _options.Margins.Left - _options.Margins.Right;
        // A print-fit pass expands percentage-positioned boxes with the layout
        // surface. Hold their fit contribution at the first resolved width;
        // in-flow boxes can still request the second bounded pass.
        double physicalContentWidth = contentWidth * (_options.PrintFitScale ?? 1D);
        double positionedOverflowCap = physicalContentWidth * 1.5D;
        if (_options.PrintFitContentWidth is double firstFitWidth) {
            positionedOverflowCap = Math.Min(positionedOverflowCap, firstFitWidth);
        }
        double overflowWidth = contentWidth;
        foreach (HtmlRenderPage page in rendered.Pages) {
            CheckCancellation();
            overflowWidth = Math.Max(overflowWidth,
                HtmlRenderScrollableOverflow.MeasureRight(page.Scene, _cancellationToken,
                    page.Margins.Left + positionedOverflowCap) - page.Margins.Left);
        }
        if (_printFitPasses < 2 && HtmlCssPrintFitResolver.TryApplyWidth(overflowWidth, _pageRules, _options)) {
            ++_printFitPasses;
            ResetLayoutPassState();
            return RenderSinglePass();
        }
        if (overflowWidth > contentWidth + 0.5D) {
            var diagnostic = new HtmlDiagnostic(ComponentName, HtmlRenderDiagnosticCodes.PrintFitOverflowUnresolved,
                "Visible layout overflow remains after bounded print fitting or exceeds the configured surface limits.",
                HtmlDiagnosticSeverity.Warning, "document",
                "content-width=" + contentWidth.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    + ";overflow-width=" + overflowWidth.ToString(System.Globalization.CultureInfo.InvariantCulture),
                OfficeConversionLossKind.Approximation);
            _diagnostics.Add(diagnostic);
            rendered = rendered.WithAdditionalDiagnostics(new[] { diagnostic });
        }
        return rendered.Project(rendered.Pages.Select(page => page.WithScene(
            HtmlRenderScrollableOverflow.RemoveLayoutBoxes(page.Scene, _cancellationToken))), rendered.Mode);
    }
}
