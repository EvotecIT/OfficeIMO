using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private IReadOnlyList<HtmlRenderPage> ApplyPageMarginContent(IReadOnlyList<HtmlRenderPage> pages) {
        var rendered = new List<HtmlRenderPage>(pages.Count);
        foreach (HtmlRenderPage page in pages) {
            var geometry = new HtmlCssPageGeometry(page.Width, page.Height, page.Margins);
            IReadOnlyDictionary<HtmlCssPageMarginPosition, HtmlCssPageMarginTemplate> boxes = _pageRules.ResolveMarginBoxes(page.PageNumber, page.PageName, geometry, _options);
            if (boxes.Count == 0) {
                rendered.Add(page);
                continue;
            }

            var visuals = new List<HtmlRenderVisual>(page.Scene);
            foreach (HtmlCssPageMarginTemplate box in boxes.Values.OrderBy(item => item.Position)) {
                string marginBoxSource = "@page @" + GetMarginBoxName(box.Position);
                if (box.Content.TryGetRunningElement(out string runningElementName, out HtmlCssRunningStringPosition runningElementPosition)) {
                    AppendRunningElementMarginBox(visuals, page, box, runningElementName, runningElementPosition);
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
                ChargeLayoutOperations(box.Content.GetRenderedLength(page.PageNumber, pages.Count, page.RunningStrings),
                    marginBoxSource + " generated content");
                string text = box.Content.Render(page.PageNumber, pages.Count, page.RunningStrings);
                double textHeight = Math.Max(1D, box.Font.Size * _options.DefaultLineHeight);
                if (text.Length == 0 || !TryGetMarginBoxBounds(page, box.Position, textHeight, out double x, out double y, out double width, out double height)) continue;
                var marginText = new HtmlRenderText(
                    text,
                    x,
                    y,
                    width,
                    height,
                    box.Font,
                    box.Color,
                    box.Alignment,
                    Math.Max(1D, box.Font.Size * _options.DefaultLineHeight),
                    _paintOrder++,
                    source: marginBoxSource,
                    semanticRole: "page-margin");
                visuals.Add(new HtmlRenderSemanticGroup(
                    HtmlRenderSemanticGroupRole.Artifact,
                    x,
                    y,
                    width,
                    height,
                    new[] { marginText },
                    _paintOrder++,
                    marginBoxSource));
            }

            rendered.Add(new HtmlRenderPage(page.PageNumber, page.Width, page.Height, visuals, page.PageName, _fonts, page.RunningStrings, page.Margins));
        }

        return rendered.AsReadOnly();
    }

    private void AppendRunningElementMarginBox(
        ICollection<HtmlRenderVisual> target,
        HtmlRenderPage page,
        HtmlCssPageMarginTemplate box,
        string name,
        HtmlCssRunningStringPosition position) {
        string snapshotValue = page.RunningStrings?.Resolve(HtmlCssRunningElementKeys.ForName(name), position) ?? string.Empty;
        if (!HtmlCssRunningElementParser.TryParseSnapshotId(snapshotValue, out int snapshotId)
            || !_runningElementSnapshots.TryGetValue(snapshotId, out HtmlCssRunningElementSnapshot? snapshot)) return;

        HtmlRenderFlowBlock block = snapshot.Block;
        if (block.Visuals.Count == 0
            || !TryGetMarginBoxBounds(page, box.Position, block.Height, out _, out _, out double availableWidth, out _)) return;

        block = ResolveRunningElementSnapshot(snapshot, page, availableWidth);
        if (block.Visuals.Count == 0
            || !TryGetMarginBoxBounds(page, box.Position, block.Height, out double x, out double y, out double width, out double height)) return;

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
        target.Add(new HtmlRenderSemanticGroup(
            HtmlRenderSemanticGroupRole.Artifact,
            x,
            y,
            width,
            height,
            marginVisuals,
            _paintOrder++,
            source));
    }

    private HtmlRenderFlowBlock ResolveRunningElementSnapshot(HtmlCssRunningElementSnapshot snapshot, HtmlRenderPage page, double availableWidth) {
        HtmlRenderFlowBlock block = snapshot.Block;
        if (!double.IsNaN(block.LayoutViewportWidth)
            && !double.IsNaN(block.LayoutViewportHeight)
            && Math.Abs(block.LayoutViewportWidth - page.Width) <= 0.0001D
            && Math.Abs(block.LayoutViewportHeight - page.Height) <= 0.0001D
            && Math.Abs(block.Width - availableWidth) <= 0.0001D) return block;

        var geometry = new HtmlCssPageGeometry(page.Width, page.Height, page.Margins);
        SetActivePageGeometry(geometry);
        HtmlRenderBoxStyle style = _styleResolver.Resolve(snapshot.Element, availableWidth, snapshot.ParentStyle);
        style.Position = "static";
        style.ZIndex = "auto";
        return LayoutElement(snapshot.Element, availableWidth, style, snapshot.ParentStyle, snapshot.Depth);
    }

    private bool TryGetMarginBoxBounds(HtmlRenderPage page, HtmlCssPageMarginPosition position, double desiredHeight, out double x, out double y, out double width, out double height) {
        desiredHeight = Math.Max(0.01D, desiredHeight);
        if (IsCorner(position)) return TryGetCornerBounds(page, position, desiredHeight, out x, out y, out width, out height);
        if (IsSide(position)) return TryGetSideBounds(page, position, desiredHeight, out x, out y, out width, out height);

        double contentWidth = Math.Max(1D, page.Width - page.Margins.Left - page.Margins.Right);
        double columnWidth = contentWidth / 3D;
        int column = position == HtmlCssPageMarginPosition.TopCenter || position == HtmlCssPageMarginPosition.BottomCenter
            ? 1
            : position == HtmlCssPageMarginPosition.TopRight || position == HtmlCssPageMarginPosition.BottomRight ? 2 : 0;
        bool top = position == HtmlCssPageMarginPosition.TopLeft || position == HtmlCssPageMarginPosition.TopCenter || position == HtmlCssPageMarginPosition.TopRight;
        double marginHeight = top ? page.Margins.Top : page.Margins.Bottom;
        if (marginHeight <= 0.01D) {
            x = y = width = height = 0D;
            return false;
        }

        x = page.Margins.Left + column * columnWidth;
        width = Math.Max(1D, columnWidth);
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
