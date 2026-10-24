using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private readonly HashSet<string> _reportedMarginTextOverflow = new HashSet<string>(StringComparer.Ordinal);

    private IReadOnlyList<HtmlRenderPage> ApplyPageMarginContent(IReadOnlyList<HtmlRenderPage> pages) {
        var rendered = new List<HtmlRenderPage>(pages.Count);
        foreach (HtmlRenderPage page in pages) {
            var geometry = new HtmlCssPageGeometry(page.Width, page.Height, page.Margins, page.PrintProduction);
            SetActivePageGeometry(geometry);
            IReadOnlyDictionary<HtmlCssPageMarginPosition, HtmlCssPageMarginTemplate> boxes = _pageRules.ResolveMarginBoxes(page.PageNumber, page.PageName, geometry, _options, _styleResolver);
            if (boxes.Count == 0) {
                rendered.Add(page);
                continue;
            }

            var visuals = new List<HtmlRenderVisual>(page.Scene);
            IReadOnlyDictionary<HtmlCssPageMarginPosition, PageMarginAxisBounds> axisBounds =
                ResolvePageMarginAxisBounds(page, boxes, pages.Count, out var preparedTextByPosition);
            foreach (HtmlCssPageMarginTemplate box in boxes.Values.OrderBy(item => GetPageMarginPaintOrder(item.Position))) {
                string marginBoxSource = "@page @" + GetMarginBoxName(box.Position);
                if (!box.Content.IsEmpty && box.UnsupportedVerticalAlignment.Length > 0
                    && _reportedMarginTextOverflow.Add(marginBoxSource + " vertical-align")) {
                    _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.GeneratedContentUnsupported,
                        "A page-margin vertical alignment used the position default.", HtmlDiagnosticSeverity.Warning,
                        marginBoxSource, "vertical-align=" + box.UnsupportedVerticalAlignment, OfficeConversionLossKind.Approximation);
                }
                if (box.Content.TryGetRunningElement(out string runningElementName, out HtmlCssRunningStringPosition runningElementPosition)) {
                    AppendRunningElementMarginBox(visuals, page, box, runningElementName, runningElementPosition, axisBounds);
                    continue;
                }
                if (box.Content.ContainsRunningElement) {
                    if (_reportedMixedRunningElementMarginBoxes.Add(marginBoxSource)) {
                        _diagnostics.Add(
                            ComponentName,
                            HtmlRenderDiagnosticCodes.GeneratedContentUnsupported,
                            "A margin-box content expression mixed element() with other generated content and was omitted.",
                            HtmlDiagnosticSeverity.Warning,
                            marginBoxSource,
                            "mixed-element-content",
                            OfficeConversionLossKind.Omission);
                    }
                    continue;
                }
                if (box.Font.Size <= 0D && !box.Typography.HasPageMarginIntrinsicHeights) {
                    AppendEmptyPageMarginBox(visuals, page, box, axisBounds);
                    continue;
                }
                string text = preparedTextByPosition.TryGetValue(box.Position, out string? preparedText)
                    ? preparedText : RenderPageMarginText(page, box, pages.Count);
                if (text.Length == 0) {
                    AppendEmptyPageMarginBox(visuals, page, box, axisBounds);
                    continue;
                }
                ResolvePageMarginTextMetrics(box, text, out double lineHeight, out double baseline,
                    out double linePaintHeight, out double paintTopOverflow);
                if (!TryGetPageMarginContentBounds(page, box, axisBounds, lineHeight, out _, out _, out double availableWidth, out _)) continue;
                // Retain the generated text and alignment in the public scene,
                // reserving enough height for its measured wrapped lines.
                HtmlInlineLayout? shaped = UsesPageMarginInlineShaping(text, box)
                    ? LayoutPageMarginShapedText(text, box, availableWidth) : null;
                IReadOnlyList<OfficeTextLine>? lines = shaped == null
                    ? MeasurePageMarginWrappedText(text, box, availableWidth).Lines : null;
                double textHeight = shaped?.Height ?? lines!.Count * lineHeight;
                if (box.Font.Size <= 0D) {
                    AppendEmptyPageMarginBox(visuals, page, box, axisBounds, textHeight);
                    continue;
                }
                if (!TryGetPageMarginContentBounds(page, box, axisBounds, textHeight, out double x, out double y, out double width, out double height)) continue;
                if (box.HasBoxGeometry) y += ResolvePageMarginVerticalOffset(box, height, textHeight);
                _fontUsage?.Observe(text, box.Font.FamilyName, box.Typography.FontDescriptor);
                IReadOnlyList<HtmlRenderVisual> marginVisuals = shaped != null
                    ? shaped.Visuals.Select(visual => visual.Translate(x, y, _paintOrder++)).ToArray()
                    : new HtmlRenderVisual[] {
                    new HtmlRenderText(text, x, y + baseline - box.Font.Size, Math.Max(0.01D, width),
                        Math.Max(0.01D, (lines!.Count - 1) * lineHeight + linePaintHeight),
                        box.Font, box.Color, box.Alignment, lineHeight, _paintOrder++,
                        linkUri: null, source: marginBoxSource, semanticRole: "page-margin", layoutY: y,
                        semanticNodeId: null, textAdvanceWidth: null,
                        layoutHeight: Math.Max(0.01D, textHeight), paintTopOverflow: paintTopOverflow,
                        ordinaryLineCarrier: true, fontDescriptor: box.Typography.FontDescriptor).WithWrappedLines(lines)
                };
                if (textHeight > height + 0.0001D || width <= 0D) {
                    // The line advance determines admission, while the full selected-face
                    // frame must survive until the shared clip owner constrains glyph paint.
                    marginVisuals = new HtmlRenderVisual[] {
                        new HtmlRenderClipGroup(x, y, width, height, width <= 0D, true,
                            marginVisuals, _paintOrder++, marginBoxSource)
                    };
                    if (_reportedMarginTextOverflow.Add(marginBoxSource)) {
                        _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.GeneratedContentUnsupported,
                            width <= 0D ? "Generated page-margin text had no available content width and was clipped."
                                : "Generated page-margin text exceeded the reserved margin height and was clipped.",
                            HtmlDiagnosticSeverity.Warning, marginBoxSource, width <= 0D ? "margin-width" : "margin-height", OfficeConversionLossKind.Omission);
                    }
                }
                AppendPageMarginVisuals(visuals, page, box, axisBounds,
                    marginVisuals, x, y, width, height, textHeight, marginBoxSource);
            }

            rendered.Add(new HtmlRenderPage(page.PageNumber, page.Width, page.Height, visuals, page.PageName, _fonts, page.RunningStrings, page.Margins, page.PrintProduction));
        }

        return rendered.AsReadOnly();
    }

    private void ResolvePageMarginTextMetrics(HtmlCssPageMarginTemplate box, string text,
        out double lineHeight, out double baseline, out double paintHeight, out double paintTopOverflow) {
        HtmlRenderBoxStyle style = box.Typography;
        InlineTextLineMetrics metrics = ResolveTextLineMetrics(text, style, allowShortLine: true);
        lineHeight = metrics.Height;
        baseline = metrics.Baseline;
        paintHeight = Math.Max(lineHeight, box.Font.Size);
        paintTopOverflow = 0D;
        if (box.Font.Size <= 0D) return;
        HtmlTextFaceMetrics? face = ResolveTextFaceMetrics(text, style);
        if (!IsUsableTextFaceMetrics(face)) return;
        paintHeight = Math.Max(lineHeight, box.Font.Size + face!.Value.Height - face.Value.BaselineOffset);
        paintTopOverflow = Math.Max(0D, face.Value.BaselineOffset - box.Font.Size);
    }

    private void AppendRunningElementMarginBox(
        ICollection<HtmlRenderVisual> target,
        HtmlRenderPage page,
        HtmlCssPageMarginTemplate box,
        string name,
        HtmlCssRunningStringPosition position,
        IReadOnlyDictionary<HtmlCssPageMarginPosition, PageMarginAxisBounds> axisBounds) {
        if (!TryGetPageMarginRunningSnapshot(page, name, position, out HtmlCssRunningElementSnapshot snapshot)) {
            AppendEmptyPageMarginBox(target, page, box, axisBounds);
            return;
        }

        HtmlRenderFlowBlock block = snapshot.Block;
        if (!TryGetPageMarginContentBounds(page, box, axisBounds, block.Height, out _, out _, out double availableWidth, out _)) return;

        block = ResolveRunningElementSnapshot(snapshot, page, availableWidth);
        if (block.Visuals.Count == 0) {
            AppendEmptyPageMarginBox(target, page, box, axisBounds, block.Height);
            return;
        }
        if (!TryGetPageMarginContentBounds(page, box, axisBounds, block.Height, out double x, out double y, out double width, out double height)) return;

        double offsetX = x;
        if (box.Alignment == OfficeTextAlignment.Center) offsetX += (width - block.Width) / 2D;
        else if (box.Alignment == OfficeTextAlignment.Right) offsetX += width - block.Width;
        double offsetY = y + ResolvePageMarginVerticalOffset(box, height, block.Height);
        var translated = block.Visuals
            .Select((visual, index) => visual.Translate(offsetX, offsetY, index))
            .ToList();
        string source = "@page @" + GetMarginBoxName(box.Position) + " element(" + name + ")";
        bool clipHorizontal = offsetX < x - 0.0001D || offsetX + block.Width > x + width + 0.0001D;
        bool clipVertical = offsetY < y - 0.0001D || offsetY + block.Height > y + height + 0.0001D;
        if (clipHorizontal || clipVertical)
            ReportPageMarginContentClip(translated, x, y, width, height, clipHorizontal, clipVertical, source);
        IReadOnlyList<HtmlRenderVisual> marginVisuals = clipHorizontal || clipVertical
            ? new HtmlRenderVisual[] {
                new HtmlRenderClipGroup(
                    x,
                    y,
                    width,
                    height,
                    clipHorizontal,
                    clipVertical,
                    translated,
                    0,
                    source)
            }
            : translated;
        AppendPageMarginVisuals(target, page, box, axisBounds, marginVisuals, x, y, width, height, block.Height, source);
    }

    private HtmlRenderFlowBlock ResolveRunningElementSnapshot(HtmlCssRunningElementSnapshot snapshot, HtmlRenderPage page, double availableWidth) {
        HtmlRenderFlowBlock block = snapshot.Block;
        if (!double.IsNaN(block.LayoutViewportWidth)
            && !double.IsNaN(block.LayoutViewportHeight)
            && Math.Abs(block.LayoutViewportWidth - page.Width) <= 0.0001D
            && Math.Abs(block.LayoutViewportHeight - page.Height) <= 0.0001D
            && Math.Abs(block.Width - availableWidth) <= 0.0001D) return block;

        var geometry = new HtmlCssPageGeometry(page.Width, page.Height, page.Margins);
        using var pageBoundaries = new DescendantPageBoundaryScope(this, retainsPageBoundaries: false);
        SetActivePageGeometry(geometry);
        HtmlRenderBoxStyle style = _styleResolver.Resolve(snapshot.Element, availableWidth, snapshot.ParentStyle);
        style.Position = "static";
        style.ZIndex = "auto";
        return LayoutElement(snapshot.Element, availableWidth, style, snapshot.ParentStyle, snapshot.Depth);
    }

    private bool TryGetMarginBoxBounds(HtmlRenderPage page, HtmlCssPageMarginPosition position, IReadOnlyDictionary<HtmlCssPageMarginPosition, PageMarginAxisBounds> axisBounds, double desiredHeight, out double x, out double y, out double width, out double height) {
        desiredHeight = Math.Max(0.01D, desiredHeight);

        bool top = position == HtmlCssPageMarginPosition.TopLeft || position == HtmlCssPageMarginPosition.TopCenter || position == HtmlCssPageMarginPosition.TopRight;
        double marginHeight = top ? page.Margins.Top : page.Margins.Bottom;
        if (marginHeight <= 0.01D || !axisBounds.TryGetValue(position, out PageMarginAxisBounds bounds)) {
            x = y = width = height = 0D;
            return false;
        }

        x = bounds.Start;
        width = bounds.Length;
        height = Math.Max(0.01D, Math.Min(desiredHeight, marginHeight));
        y = top
            ? Math.Max(0D, (marginHeight - height) / 2D)
            : page.Height - marginHeight + Math.Max(0D, (marginHeight - height) / 2D);
        return true;
    }

    private static double ResolvePageMarginVerticalOffset(HtmlCssPageMarginTemplate box, double available, double content) =>
        box.VerticalAlignment switch {
            OfficeTextVerticalAlignment.Bottom => Math.Max(0D, available - content),
            OfficeTextVerticalAlignment.Center => Math.Max(0D, (available - content) / 2D),
            _ => 0D
        };

    private static bool IsLeftMargin(HtmlCssPageMarginPosition position) => position is
        HtmlCssPageMarginPosition.LeftTop or HtmlCssPageMarginPosition.LeftMiddle or HtmlCssPageMarginPosition.LeftBottom
        or HtmlCssPageMarginPosition.TopLeftCorner or HtmlCssPageMarginPosition.BottomLeftCorner;

    private static bool IsTopMargin(HtmlCssPageMarginPosition position) => position is
        HtmlCssPageMarginPosition.TopLeft or HtmlCssPageMarginPosition.TopCenter or HtmlCssPageMarginPosition.TopRight
        or HtmlCssPageMarginPosition.TopLeftCorner or HtmlCssPageMarginPosition.TopRightCorner;

    private static bool IsCorner(HtmlCssPageMarginPosition position) =>
        position == HtmlCssPageMarginPosition.TopLeftCorner
        || position == HtmlCssPageMarginPosition.TopRightCorner
        || position == HtmlCssPageMarginPosition.BottomLeftCorner
        || position == HtmlCssPageMarginPosition.BottomRightCorner;

    private static bool IsSide(HtmlCssPageMarginPosition position) =>
        position == HtmlCssPageMarginPosition.LeftTop
        || position == HtmlCssPageMarginPosition.LeftMiddle
        || position == HtmlCssPageMarginPosition.LeftBottom
        || position == HtmlCssPageMarginPosition.RightTop
        || position == HtmlCssPageMarginPosition.RightMiddle
        || position == HtmlCssPageMarginPosition.RightBottom;

    private static string GetMarginBoxName(HtmlCssPageMarginPosition position) =>
        HtmlCssPageSettingsResolver.GetMarginBoxName(position);
}
