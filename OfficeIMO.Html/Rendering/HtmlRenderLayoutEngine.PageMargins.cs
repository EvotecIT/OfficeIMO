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
            IReadOnlyDictionary<HtmlCssPageMarginPosition, PageMarginHorizontalBounds> horizontalBounds =
                ResolveHorizontalPageMarginBounds(page, boxes, pages.Count, out var horizontalText);
            foreach (HtmlCssPageMarginTemplate box in boxes.Values.OrderBy(item => item.Position)) {
                string marginBoxSource = "@page @" + GetMarginBoxName(box.Position);
                if (box.Content.TryGetRunningElement(out string runningElementName, out HtmlCssRunningStringPosition runningElementPosition)) {
                    AppendRunningElementMarginBox(visuals, page, box, runningElementName, runningElementPosition, horizontalBounds);
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
                if (box.Font.Size <= 0D) {
                    AppendEmptyPageMarginBox(visuals, page, box, horizontalBounds);
                    continue;
                }
                string text = horizontalText.TryGetValue(box.Position, out string? preparedText)
                    ? preparedText : RenderPageMarginText(page, box, pages.Count);
                if (text.Length == 0) {
                    AppendEmptyPageMarginBox(visuals, page, box, horizontalBounds);
                    continue;
                }
                ResolvePageMarginTextMetrics(box, text, out double lineHeight, out double baseline,
                    out double linePaintHeight, out double paintTopOverflow);
                if (!TryGetPageMarginContentBounds(page, box, horizontalBounds, lineHeight, out _, out _, out double availableWidth, out _)) continue;
                // Retain the generated text and alignment in the public scene,
                // reserving enough height for its measured wrapped lines.
                OfficeTextBlockLayout measured = OfficeTextLayoutEngine.LayoutTextBlock(text, box.Font.Size,
                    Math.Max(0.01D, availableWidth), double.MaxValue, 1D, box.Font.Size,
                    (value, _) => MeasurePageMarginText(value ?? string.Empty, box),
                    wrap: true, forceSingleLine: false, shrinkToFit: false,
                    overflowBehavior: OfficeTextOverflowBehavior.Clip);
                IReadOnlyList<OfficeTextLine> lines = measured.Lines;
                if (measured.Clipped && _reportedMarginTextOverflow.Add(marginBoxSource + " text-layout-limit")) {
                    _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.GeneratedContentUnsupported,
                        "Generated page-margin text exceeded the shared text-layout limits and was truncated.",
                        HtmlDiagnosticSeverity.Warning, marginBoxSource, "text-layout-limit", OfficeConversionLossKind.Omission);
                }
                double textHeight = lines.Count * lineHeight;
                if (!TryGetPageMarginContentBounds(page, box, horizontalBounds, textHeight, out double x, out double y, out double width, out double height)) continue;
                if (box.HasBoxGeometry) y += Math.Max(0D, (height - textHeight) / 2D);
                _fontUsage?.Observe(text, box.Font.FamilyName, box.Typography.FontDescriptor);
                IReadOnlyList<HtmlRenderVisual> marginVisuals = new HtmlRenderVisual[] {
                    new HtmlRenderText(text, x, y + baseline - box.Font.Size, Math.Max(0.01D, width),
                        Math.Max(0.01D, (lines.Count - 1) * lineHeight + linePaintHeight),
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
                AppendPageMarginVisuals(visuals, page, box, horizontalBounds,
                    marginVisuals, x, y, width, height, marginBoxSource);
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
        IReadOnlyDictionary<HtmlCssPageMarginPosition, PageMarginHorizontalBounds> horizontalBounds) {
        if (!TryGetPageMarginRunningSnapshot(page, name, position, out HtmlCssRunningElementSnapshot snapshot)) {
            AppendEmptyPageMarginBox(target, page, box, horizontalBounds);
            return;
        }

        HtmlRenderFlowBlock block = snapshot.Block;
        if (block.Visuals.Count == 0) {
            AppendEmptyPageMarginBox(target, page, box, horizontalBounds);
            return;
        }
        if (!TryGetPageMarginContentBounds(page, box, horizontalBounds, block.Height, out _, out _, out double availableWidth, out _)) return;

        block = ResolveRunningElementSnapshot(snapshot, page, availableWidth);
        if (block.Visuals.Count == 0) {
            AppendEmptyPageMarginBox(target, page, box, horizontalBounds);
            return;
        }
        if (!TryGetPageMarginContentBounds(page, box, horizontalBounds, block.Height, out double x, out double y, out double width, out double height)) return;

        double offsetX = x;
        if (box.Alignment == OfficeTextAlignment.Center) offsetX += (width - block.Width) / 2D;
        else if (box.Alignment == OfficeTextAlignment.Right) offsetX += width - block.Width;
        double offsetY = y + (height - block.Height) / 2D;
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
        AppendPageMarginVisuals(target, page, box, horizontalBounds, marginVisuals, x, y, width, height, source);
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

    private bool TryGetMarginBoxBounds(HtmlRenderPage page, HtmlCssPageMarginPosition position, IReadOnlyDictionary<HtmlCssPageMarginPosition, PageMarginHorizontalBounds> horizontalBounds, double desiredHeight, out double x, out double y, out double width, out double height) {
        desiredHeight = Math.Max(0.01D, desiredHeight);
        if (IsCorner(position)) return TryGetCornerBounds(page, position, desiredHeight, out x, out y, out width, out height);
        if (IsSide(position)) return TryGetSideBounds(page, position, desiredHeight, out x, out y, out width, out height);

        bool top = position == HtmlCssPageMarginPosition.TopLeft || position == HtmlCssPageMarginPosition.TopCenter || position == HtmlCssPageMarginPosition.TopRight;
        double marginHeight = top ? page.Margins.Top : page.Margins.Bottom;
        if (marginHeight <= 0.01D || !horizontalBounds.TryGetValue(position, out PageMarginHorizontalBounds bounds)) {
            x = y = width = height = 0D;
            return false;
        }

        x = bounds.X;
        width = bounds.Width;
        height = Math.Max(0.01D, Math.Min(desiredHeight, marginHeight));
        y = top
            ? Math.Max(0D, (marginHeight - height) / 2D)
            : page.Height - marginHeight + Math.Max(0D, (marginHeight - height) / 2D);
        return true;
    }

    private bool TryGetCornerBounds(HtmlRenderPage page, HtmlCssPageMarginPosition position, double lineHeight, out double x, out double y, out double width, out double height) {
        bool left = position == HtmlCssPageMarginPosition.TopLeftCorner || position == HtmlCssPageMarginPosition.BottomLeftCorner;
        bool top = position == HtmlCssPageMarginPosition.TopLeftCorner || position == HtmlCssPageMarginPosition.TopRightCorner;
        double marginWidth = left ? page.Margins.Left : page.Margins.Right;
        double marginHeight = top ? page.Margins.Top : page.Margins.Bottom;
        if (marginWidth <= 0.01D || marginHeight <= 0.01D) {
            x = y = width = height = 0D;
            return false;
        }

        x = left ? 0D : page.Width - marginWidth;
        width = marginWidth;
        height = Math.Max(0.01D, Math.Min(lineHeight, marginHeight));
        y = top
            ? Math.Max(0D, (marginHeight - height) / 2D)
            : page.Height - marginHeight + Math.Max(0D, (marginHeight - height) / 2D);
        return true;
    }

    private bool TryGetSideBounds(HtmlRenderPage page, HtmlCssPageMarginPosition position, double lineHeight, out double x, out double y, out double width, out double height) {
        bool left = position == HtmlCssPageMarginPosition.LeftTop || position == HtmlCssPageMarginPosition.LeftMiddle || position == HtmlCssPageMarginPosition.LeftBottom;
        double marginWidth = left ? page.Margins.Left : page.Margins.Right;
        double contentHeight = Math.Max(1D, page.Height - page.Margins.Top - page.Margins.Bottom);
        if (marginWidth <= 0.01D) {
            x = y = width = height = 0D;
            return false;
        }

        int section = position == HtmlCssPageMarginPosition.LeftMiddle || position == HtmlCssPageMarginPosition.RightMiddle
            ? 1
            : position == HtmlCssPageMarginPosition.LeftBottom || position == HtmlCssPageMarginPosition.RightBottom ? 2 : 0;
        double sectionHeight = contentHeight / 3D;
        x = left ? 0D : page.Width - marginWidth;
        width = marginWidth;
        height = Math.Max(0.01D, Math.Min(lineHeight, sectionHeight));
        y = page.Margins.Top + section * sectionHeight + Math.Max(0D, (sectionHeight - height) / 2D);
        return true;
    }

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

    private static string GetMarginBoxName(HtmlCssPageMarginPosition position) {
        switch (position) {
            case HtmlCssPageMarginPosition.TopLeftCorner: return "top-left-corner";
            case HtmlCssPageMarginPosition.TopLeft: return "top-left";
            case HtmlCssPageMarginPosition.TopCenter: return "top-center";
            case HtmlCssPageMarginPosition.TopRight: return "top-right";
            case HtmlCssPageMarginPosition.TopRightCorner: return "top-right-corner";
            case HtmlCssPageMarginPosition.LeftTop: return "left-top";
            case HtmlCssPageMarginPosition.LeftMiddle: return "left-middle";
            case HtmlCssPageMarginPosition.LeftBottom: return "left-bottom";
            case HtmlCssPageMarginPosition.RightTop: return "right-top";
            case HtmlCssPageMarginPosition.RightMiddle: return "right-middle";
            case HtmlCssPageMarginPosition.RightBottom: return "right-bottom";
            case HtmlCssPageMarginPosition.BottomLeftCorner: return "bottom-left-corner";
            case HtmlCssPageMarginPosition.BottomLeft: return "bottom-left";
            case HtmlCssPageMarginPosition.BottomCenter: return "bottom-center";
            case HtmlCssPageMarginPosition.BottomRight: return "bottom-right";
            default: return "bottom-right-corner";
        }
    }
}
