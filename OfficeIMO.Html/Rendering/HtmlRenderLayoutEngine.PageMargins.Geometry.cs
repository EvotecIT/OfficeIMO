namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>Returns the inner frame of a horizontal generated box while preserving the legacy text-only route.</summary>
    private bool TryGetPageMarginContentBounds(HtmlRenderPage page, HtmlCssPageMarginTemplate box,
        IReadOnlyDictionary<HtmlCssPageMarginPosition, PageMarginHorizontalBounds> horizontalBounds,
        double desiredHeight, out double x, out double y, out double width, out double height) {
        if (!box.HasBoxGeometry) return TryGetMarginBoxBounds(page, box.Position, horizontalBounds,
            desiredHeight, out x, out y, out width, out height);
        if (!TryGetPageMarginBorderBounds(page, box, horizontalBounds, desiredHeight, out x, out y, out width, out height)) return false;
        HtmlRenderBoxStyle style = box.Typography;
        x += style.BorderLeftWidth + style.PaddingLeft;
        y += style.BorderTopWidth + style.PaddingTop;
        width = Math.Max(0D, width - style.HorizontalInsets);
        height = Math.Max(0D, height - style.VerticalInsets);
        return true;
    }

    /// <summary>Resolves the border frame inside signed margins without allowing negative content dimensions.</summary>
    private bool TryGetPageMarginBorderBounds(HtmlRenderPage page, HtmlCssPageMarginTemplate box,
        IReadOnlyDictionary<HtmlCssPageMarginPosition, PageMarginHorizontalBounds> horizontalBounds,
        double naturalContentHeight, out double x, out double y, out double width, out double height) {
        CheckCancellation();
        bool top = box.Position is HtmlCssPageMarginPosition.TopLeft or HtmlCssPageMarginPosition.TopCenter or HtmlCssPageMarginPosition.TopRight;
        double reservedHeight = Math.Max(0D, top ? page.Margins.Top : page.Margins.Bottom);
        if (!horizontalBounds.TryGetValue(box.Position, out PageMarginHorizontalBounds bounds)) {
            x = y = width = height = 0D;
            return false;
        }
        HtmlRenderBoxStyle style = box.Typography;
        x = bounds.X + style.MarginLeft;
        width = Math.Max(style.HorizontalInsets, bounds.Width - style.MarginLeft - style.MarginRight);
        double autoContentHeight = Math.Max(0D, reservedHeight - style.VerticalInsets
            - (style.MarginTopAuto ? 0D : style.MarginTop) - (style.MarginBottomAuto ? 0D : style.MarginBottom));
        // Intrinsic block sizes use content after layout at the final assigned
        // inner width. Convert to authored box coordinates before the common
        // numeric constraint owner adds insets and applies minimum precedence.
        double intrinsicHeight = naturalContentHeight + (style.BorderBox ? style.VerticalInsets : 0D);
        height = !style.HasPageMarginIntrinsicHeights && !style.ExplicitHeight.HasValue && !style.MinHeight.HasValue && !style.MaxHeight.HasValue
            && style.MarginTop == 0D && style.MarginBottom == 0D
            ? Math.Max(reservedHeight, style.VerticalInsets)
            : ResolveBoxHeight(autoContentHeight, width, style,
                style.PageMarginIntrinsicHeight ? intrinsicHeight : style.ExplicitHeight,
                style.PageMarginIntrinsicMinHeight ? intrinsicHeight : style.MinHeight,
                style.PageMarginIntrinsicMaxHeight ? intrinsicHeight : style.MaxHeight, allowZeroHeight: true);
        double marginTop = ResolvePageMarginTopMargin(style, reservedHeight, height, top);
        y = (top ? 0D : page.Height - reservedHeight) + marginTop;
        if (!Finite(x) || !Finite(y) || !Finite(width) || !Finite(height) || !Finite(x + width) || !Finite(y + height)) {
            string source = "@page @" + GetMarginBoxName(box.Position);
            if (_reportedMarginTextOverflow.Add(source + " geometry")) {
                _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.GeneratedContentUnsupported,
                    "Generated page-margin geometry exceeded finite layout coordinates and the box was omitted.",
                    HtmlDiagnosticSeverity.Warning, source, "margin-geometry", OfficeConversionLossKind.Omission);
            }
            x = y = width = height = 0D;
            return false;
        }
        return true;

        static bool Finite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
    }

    /// <summary>Balances the fixed-axis equation after height constraints, preserving the body-facing margin.</summary>
    private static double ResolvePageMarginTopMargin(HtmlRenderBoxStyle style, double stripHeight, double borderHeight, bool top) {
        bool topAuto = style.MarginTopAuto, bottomAuto = style.MarginBottomAuto;
        double marginTop = topAuto ? 0D : style.MarginTop;
        double marginBottom = bottomAuto ? 0D : style.MarginBottom;
        if (borderHeight + marginTop + marginBottom > stripHeight) topAuto = bottomAuto = false;
        // With no remaining auto value, the outer-facing margin absorbs the
        // difference. This also preserves the existing body-edge inset floor.
        if (!topAuto && !bottomAuto) {
            if (top) topAuto = true;
            else bottomAuto = true;
        }
        if (!topAuto) return marginTop;
        return bottomAuto ? (stripHeight - borderHeight) / 2D : stripHeight - borderHeight - marginBottom;
    }

    private void AppendEmptyPageMarginBox(ICollection<HtmlRenderVisual> target, HtmlRenderPage page,
        HtmlCssPageMarginTemplate box, IReadOnlyDictionary<HtmlCssPageMarginPosition, PageMarginHorizontalBounds> bounds,
        double naturalContentHeight = 0D) {
        if (!box.HasBoxGeometry || box.Content.IsEmpty) return;
        if (!TryGetPageMarginBorderBounds(page, box, bounds, naturalContentHeight, out double x, out double y, out double width, out double height)) return;
        AppendPageMarginVisuals(target, page, box, bounds, Array.Empty<HtmlRenderVisual>(), x, y, width, height,
            naturalContentHeight, "@page @" + GetMarginBoxName(box.Position));
    }

    private void AppendPageMarginVisuals(ICollection<HtmlRenderVisual> target, HtmlRenderPage page,
        HtmlCssPageMarginTemplate box, IReadOnlyDictionary<HtmlCssPageMarginPosition, PageMarginHorizontalBounds> bounds,
        IReadOnlyList<HtmlRenderVisual> content, double x, double y, double width, double height,
        double naturalContentHeight, string source) {
        IReadOnlyList<HtmlRenderVisual> children = content;
        if (box.HasBoxGeometry) {
            double contentWidth = width, contentHeight = height;
            if (!TryGetPageMarginBorderBounds(page, box, bounds, naturalContentHeight, out x, out y, out width, out height)) return;
            var decorated = new List<HtmlRenderVisual>();
            AddBoxBackgroundCore(decorated, box.Typography, x, y, width, height,
                box.Typography.BorderInsets, default, source, source, source);
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
