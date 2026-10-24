using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>Uses the same effective-provider admission as ordinary inline text.</summary>
    private bool UsesPageMarginInlineShaping(string text, HtmlCssPageMarginTemplate box) =>
        box.Font.Size > 0D && (_options.TextShapingProvider != null
            || !box.Typography.TextFeatureSettings.IsDefault || OfficeManagedTextShaper.RequiresComplexLayout(text));

    /// <summary>
    /// Admits generated hard lines through Core, then lets the ordinary inline
    /// owner retain font fallback, bidi, complete clusters and positioned glyphs.
    /// OfficeTextLine itself cannot carry the source offsets of prepared glyphs.
    /// </summary>
    private List<HtmlInlineRun> CreatePageMarginTextRuns(string text, HtmlCssPageMarginTemplate box,
        out HtmlRenderBoxStyle style) {
        CheckCancellation();
        string source = "@page @" + GetMarginBoxName(box.Position);
        ChargeLayoutOperations(text.Length, source + " text preparation");
        OfficeTextBlockLayout admitted = OfficeTextLayoutEngine.LayoutTextBlock(text, box.Font.Size,
            double.MaxValue, double.MaxValue, 1D, box.Font.Size, static (_, _) => 0D,
            wrap: true, forceSingleLine: false, shrinkToFit: false, overflowBehavior: OfficeTextOverflowBehavior.Clip);
        if (admitted.Clipped) ReportPageMarginTextLayoutLimit(box);
        style = box.Typography.Clone();
        style.Alignment = box.Alignment;
        style.SemanticRole = "page-margin";
        style.TextIndent = null;
        // Generated content already owns its literal whitespace and hard lines.
        // Core above expands tabs with the same rule as the plain margin carrier.
        style.PreserveWhitespace = true;
        style.BreakSpaces = false;
        var sourceRuns = new List<HtmlInlineRun>();
        for (int index = 0; index < admitted.Lines.Count; index++) {
            CheckCancellation();
            if (index > 0) sourceRuns.Add(new HtmlInlineRun("\n", style, null, source));
            string line = admitted.Lines[index].Text;
            if (line.Length == 0) continue;
            ReportUnsupportedComplexTextShaping(line, null, style, source);
            sourceRuns.Add(new HtmlInlineRun(line, style, null, source));
        }
        // A final forced break commits the final authored line even when it is
        // empty; ordinary inline flow otherwise drops a trailing empty line.
        sourceRuns.Add(new HtmlInlineRun("\u2028", style, null, source));
        return ApplyScopedFontFallbacks(sourceRuns);
    }

    private PageMarginIntrinsicSize MeasurePageMarginShapedTextIntrinsics(string text, HtmlCssPageMarginTemplate box) {
        List<HtmlInlineRun> inline = CreatePageMarginTextRuns(text, box, out _);
        var intrinsic = new List<IntrinsicTextRun>();
        foreach (HtmlInlineRun run in inline) {
            CheckCancellation();
            intrinsic.Add(run.Text is "\n" or "\u2028" ? IntrinsicTextRun.ForcedBreak(run.Style)
                : new IntrinsicTextRun(run.Text, run.Style, textTransformPending: false) {
                    PreparedShaping = run.PreparedShaping,
                    PreserveContextualFallback = run.PreserveContextualFallback
                });
        }
        double maximum = MeasureMaxContentRuns(intrinsic);
        return new PageMarginIntrinsicSize(true, Math.Min(maximum, MeasureMinContentRuns(intrinsic)), maximum);
    }

    private HtmlInlineLayout LayoutPageMarginShapedText(string text, HtmlCssPageMarginTemplate box, double width) {
        List<HtmlInlineRun> runs = CreatePageMarginTextRuns(text, box, out HtmlRenderBoxStyle style);
        // Margin text has always broken overlong words. The ordinary emergency
        // break owner slices accepted clusters; it never cuts a retained ligature.
        foreach (HtmlInlineRun run in runs) run.Style.OverflowWrap = "break-word";
        style.OverflowWrap = "break-word";
        return LayoutInlineRuns(runs, Math.Max(0.01D, width), style,
            maximumLines: OfficeTextLayoutEngine.MaximumLayoutLines,
            onLineLimitExceeded: () => ReportPageMarginTextLayoutLimit(box), retainPreparedTokenClusters: true);
    }

    private void ReportPageMarginTextLayoutLimit(HtmlCssPageMarginTemplate box) {
        string source = "@page @" + GetMarginBoxName(box.Position);
        if (!_reportedMarginTextOverflow.Add(source + " text-layout-limit")) return;
        _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.GeneratedContentUnsupported,
            "Generated page-margin text exceeded the shared text-layout limits and was truncated.",
            HtmlDiagnosticSeverity.Warning, source, "text-layout-limit", OfficeConversionLossKind.Omission);
    }
}
