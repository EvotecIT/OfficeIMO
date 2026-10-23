namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>Returns the inner frame of a generated box while preserving the legacy text-only route.</summary>
    private bool TryGetPageMarginContentBounds(HtmlRenderPage page, HtmlCssPageMarginTemplate box,
        IReadOnlyDictionary<HtmlCssPageMarginPosition, PageMarginAxisBounds> axisBounds,
        double desiredHeight, out double x, out double y, out double width, out double height) {
        if (!box.HasBoxGeometry) return TryGetMarginBoxBounds(page, box.Position, axisBounds,
            desiredHeight, out x, out y, out width, out height);
        if (!TryGetPageMarginBorderBounds(page, box, axisBounds, desiredHeight, out x, out y, out width, out height)) return false;
        HtmlRenderBoxStyle style = box.Typography;
        x += style.BorderLeftWidth + style.PaddingLeft;
        y += style.BorderTopWidth + style.PaddingTop;
        width = Math.Max(0D, width - style.HorizontalInsets);
        height = Math.Max(0D, height - style.VerticalInsets);
        return true;
    }

    /// <summary>Resolves the border frame inside signed margins without allowing negative content dimensions.</summary>
    private bool TryGetPageMarginBorderBounds(HtmlRenderPage page, HtmlCssPageMarginTemplate box,
        IReadOnlyDictionary<HtmlCssPageMarginPosition, PageMarginAxisBounds> axisBounds,
        double naturalContentHeight, out double x, out double y, out double width, out double height) {
        CheckCancellation();
        bool side = IsSide(box.Position), corner = IsCorner(box.Position);
        bool top = IsTopMargin(box.Position);
        bool left = IsLeftMargin(box.Position);
        double reservedHeight = Math.Max(0D, top ? page.Margins.Top : page.Margins.Bottom);
        if (!axisBounds.TryGetValue(box.Position, out PageMarginAxisBounds bounds)) {
            x = y = width = height = 0D;
            return false;
        }
        HtmlRenderBoxStyle style = box.Typography;
        if (side || corner) {
            double reservedWidth = Math.Max(0D, left ? page.Margins.Left : page.Margins.Right);
            width = bounds.FixedWidth ?? 0D;
            x = (left ? 0D : page.Width - reservedWidth) + ResolvePageMarginFixedStart(reservedWidth, width,
                style.MarginLeft, style.MarginRight, style.MarginLeftAuto, style.MarginRightAuto, left);
        } else {
            x = bounds.Start + (style.MarginLeftAuto ? 0D : style.MarginLeft);
            width = Math.Max(style.HorizontalInsets, bounds.Length
                - (style.MarginLeftAuto ? 0D : style.MarginLeft) - (style.MarginRightAuto ? 0D : style.MarginRight));
        }
        if (side) {
            y = bounds.Start + (style.MarginTopAuto ? 0D : style.MarginTop);
            height = Math.Max(style.VerticalInsets, bounds.Length
                - (style.MarginTopAuto ? 0D : style.MarginTop) - (style.MarginBottomAuto ? 0D : style.MarginBottom));
        } else {
            double autoContentHeight = Math.Max(0D, reservedHeight - style.VerticalInsets
                - (style.MarginTopAuto ? 0D : style.MarginTop) - (style.MarginBottomAuto ? 0D : style.MarginBottom));
            // Intrinsic block sizes use content laid out at the final inner width.
            double intrinsicHeight = naturalContentHeight + (style.BorderBox ? style.VerticalInsets : 0D);
            height = ResolveBoxHeight(autoContentHeight, width, style,
                style.PageMarginIntrinsicHeight ? intrinsicHeight : style.ExplicitHeight,
                style.PageMarginIntrinsicMinHeight ? intrinsicHeight : style.MinHeight,
                style.PageMarginIntrinsicMaxHeight ? intrinsicHeight : style.MaxHeight, allowZeroHeight: true);
            y = (top ? 0D : page.Height - reservedHeight) + ResolvePageMarginFixedStart(reservedHeight, height,
                style.MarginTop, style.MarginBottom, style.MarginTopAuto, style.MarginBottomAuto, top);
        }
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

    /// <summary>Resolves fixed-axis widths with shared intrinsic values and numeric box constraints.</summary>
    private static double ResolvePageMarginFixedWidth(HtmlRenderBoxStyle style, PageMarginIntrinsicSize intrinsic, double availableWidth) {
        style = ApplyIntrinsicWidthValues(style, intrinsic.Minimum, intrinsic.Maximum, availableWidth);
        double available = availableWidth - (style.MarginLeftAuto ? 0D : style.MarginLeft)
            - (style.MarginRightAuto ? 0D : style.MarginRight);
        return ResolveBoxWidth(available, style, allowZeroWidth: true);
    }

    /// <summary>Balances either fixed axis, preserving the body-facing margin when overconstrained.</summary>
    private static double ResolvePageMarginFixedStart(double available, double borderExtent,
        double leading, double trailing, bool leadingAuto, bool trailingAuto, bool outerLeading) {
        double before = leadingAuto ? 0D : leading;
        double after = trailingAuto ? 0D : trailing;
        if (borderExtent + before + after > available) leadingAuto = trailingAuto = false;
        if (!leadingAuto && !trailingAuto) {
            if (outerLeading) leadingAuto = true;
            else trailingAuto = true;
        }
        if (!leadingAuto) return before;
        return trailingAuto ? (available - borderExtent) / 2D : available - borderExtent - after;
    }

    private void AppendEmptyPageMarginBox(ICollection<HtmlRenderVisual> target, HtmlRenderPage page,
        HtmlCssPageMarginTemplate box, IReadOnlyDictionary<HtmlCssPageMarginPosition, PageMarginAxisBounds> bounds,
        double naturalContentHeight = 0D) {
        if (!box.HasBoxGeometry || box.Content.IsEmpty) return;
        if (!TryGetPageMarginBorderBounds(page, box, bounds, naturalContentHeight, out double x, out double y, out double width, out double height)) return;
        AppendPageMarginVisuals(target, page, box, bounds, Array.Empty<HtmlRenderVisual>(), x, y, width, height,
            naturalContentHeight, "@page @" + GetMarginBoxName(box.Position));
    }

    private void AppendPageMarginVisuals(ICollection<HtmlRenderVisual> target, HtmlRenderPage page,
        HtmlCssPageMarginTemplate box, IReadOnlyDictionary<HtmlCssPageMarginPosition, PageMarginAxisBounds> bounds,
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
