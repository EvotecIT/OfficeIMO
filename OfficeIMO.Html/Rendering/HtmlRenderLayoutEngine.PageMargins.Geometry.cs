namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>Returns the inner frame of a horizontal generated box while preserving the legacy text-only route.</summary>
    private bool TryGetPageMarginContentBounds(HtmlRenderPage page, HtmlCssPageMarginTemplate box,
        IReadOnlyDictionary<HtmlCssPageMarginPosition, PageMarginHorizontalBounds> horizontalBounds,
        double desiredHeight, out double x, out double y, out double width, out double height) {
        if (!box.HasBoxGeometry) return TryGetMarginBoxBounds(page, box.Position, horizontalBounds,
            desiredHeight, out x, out y, out width, out height);
        if (!TryGetPageMarginBorderBounds(page, box, horizontalBounds, out x, out y, out width, out height)) return false;
        HtmlRenderBoxStyle style = box.Typography;
        x += style.BorderLeftWidth + style.PaddingLeft;
        y += style.BorderTopWidth + style.PaddingTop;
        width = Math.Max(0D, width - style.HorizontalInsets);
        height = Math.Max(0D, height - style.VerticalInsets);
        return true;
    }

    private static bool TryGetPageMarginBorderBounds(HtmlRenderPage page, HtmlCssPageMarginTemplate box,
        IReadOnlyDictionary<HtmlCssPageMarginPosition, PageMarginHorizontalBounds> horizontalBounds,
        out double x, out double y, out double width, out double height) {
        bool top = box.Position is HtmlCssPageMarginPosition.TopLeft or HtmlCssPageMarginPosition.TopCenter or HtmlCssPageMarginPosition.TopRight;
        double reservedHeight = Math.Max(0D, top ? page.Margins.Top : page.Margins.Bottom);
        if (!horizontalBounds.TryGetValue(box.Position, out PageMarginHorizontalBounds bounds)) {
            x = y = width = height = 0D;
            return false;
        }
        // Auto content height cannot become negative when padding and borders
        // exceed the strip. The inside edge stays attached to the page body.
        height = Math.Max(reservedHeight, box.Typography.VerticalInsets);
        x = bounds.X;
        y = top ? reservedHeight - height : page.Height - reservedHeight;
        width = bounds.Width;
        return true;
    }

    private void AppendEmptyPageMarginBox(ICollection<HtmlRenderVisual> target, HtmlRenderPage page,
        HtmlCssPageMarginTemplate box, IReadOnlyDictionary<HtmlCssPageMarginPosition, PageMarginHorizontalBounds> bounds) {
        if (!box.HasBoxGeometry || box.Content.IsEmpty) return;
        if (!TryGetPageMarginBorderBounds(page, box, bounds, out double x, out double y, out double width, out double height)) return;
        AppendPageMarginVisuals(target, page, box, bounds, Array.Empty<HtmlRenderVisual>(), x, y, width, height,
            "@page @" + GetMarginBoxName(box.Position));
    }

    private void AppendPageMarginVisuals(ICollection<HtmlRenderVisual> target, HtmlRenderPage page,
        HtmlCssPageMarginTemplate box, IReadOnlyDictionary<HtmlCssPageMarginPosition, PageMarginHorizontalBounds> bounds,
        IReadOnlyList<HtmlRenderVisual> content, double x, double y, double width, double height, string source) {
        IReadOnlyList<HtmlRenderVisual> children = content;
        if (box.HasBoxGeometry) {
            double contentWidth = width, contentHeight = height;
            if (!TryGetPageMarginBorderBounds(page, box, bounds, out x, out y, out width, out height)) return;
            var decorated = new List<HtmlRenderVisual>();
            if (box.Typography.UnsupportedBorderPaint.Length > 0) {
                if (_reportedBorderPaintFallbacks.Add(source)) {
                    _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.BorderPaintValueUnsupported,
                        "A page-margin CSS border declaration used no-border fallback.", HtmlDiagnosticSeverity.Warning,
                        source, box.Typography.UnsupportedBorderPaint, OfficeConversionLossKind.Omission);
                }
            } else {
                AddBorderPaintCore(decorated, box.Typography, x, y, width, height, default, source);
            }
            decorated.AddRange(content);
            children = decorated;
            ReportPageMarginPhysicalOverflow(children, page, source);
            if ((contentWidth <= 0D || contentHeight <= 0D) && content.Count > 0 && _reportedMarginTextOverflow.Add(source)) {
                _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.GeneratedContentUnsupported,
                    "Generated page-margin content had no available content area and was clipped.",
                    HtmlDiagnosticSeverity.Warning, source, contentWidth <= 0D ? "margin-width" : "margin-height", OfficeConversionLossKind.Omission);
            }
        }
        // Paint-neutral containers need a positive carrier; the child clip keeps
        // an exact zero content height and cannot acquire paint from this floor.
        target.Add(new HtmlRenderSemanticGroup(HtmlRenderSemanticGroupRole.Artifact,
            x, y, width, Math.Max(0.01D, height), children, _paintOrder++, source));
    }
}
