using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private IReadOnlyList<HtmlRenderPage> ApplyPageMarginContent(IReadOnlyList<HtmlRenderPage> pages) {
        var rendered = new List<HtmlRenderPage>(pages.Count);
        foreach (HtmlRenderPage page in pages) {
            IReadOnlyDictionary<HtmlCssPageMarginPosition, HtmlCssPageMarginTemplate> boxes = _pageRules.ResolveMarginBoxes(page.PageNumber, page.PageName);
            if (boxes.Count == 0) {
                rendered.Add(page);
                continue;
            }

            var visuals = new List<HtmlRenderVisual>(page.Scene);
            foreach (HtmlCssPageMarginTemplate box in boxes.Values.OrderBy(item => item.Position)) {
                ChargeLayoutOperations(box.Content.GetRenderedLength(page.PageNumber, pages.Count),
                    "@page @" + GetMarginBoxName(box.Position) + " generated content");
                string text = box.Content.Render(page.PageNumber, pages.Count);
                if (text.Length == 0 || !TryGetMarginBoxBounds(page, box.Position, box.Font.Size, out double x, out double y, out double width, out double height)) continue;
                visuals.Add(new HtmlRenderText(
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
                    source: "@page @" + GetMarginBoxName(box.Position),
                    semanticRole: "page-margin"));
            }

            rendered.Add(new HtmlRenderPage(page.PageNumber, page.Width, page.Height, visuals, page.PageName, _fonts));
        }

        return rendered.AsReadOnly();
    }

    private bool TryGetMarginBoxBounds(HtmlRenderPage page, HtmlCssPageMarginPosition position, double fontSize, out double x, out double y, out double width, out double height) {
        double lineHeight = Math.Max(1D, fontSize * _options.DefaultLineHeight);
        if (IsCorner(position)) return TryGetCornerBounds(page, position, lineHeight, out x, out y, out width, out height);
        if (IsSide(position)) return TryGetSideBounds(page, position, lineHeight, out x, out y, out width, out height);

        double contentWidth = Math.Max(1D, page.Width - _options.Margins.Left - _options.Margins.Right);
        double columnWidth = contentWidth / 3D;
        int column = position == HtmlCssPageMarginPosition.TopCenter || position == HtmlCssPageMarginPosition.BottomCenter
            ? 1
            : position == HtmlCssPageMarginPosition.TopRight || position == HtmlCssPageMarginPosition.BottomRight ? 2 : 0;
        bool top = position == HtmlCssPageMarginPosition.TopLeft || position == HtmlCssPageMarginPosition.TopCenter || position == HtmlCssPageMarginPosition.TopRight;
        double marginHeight = top ? _options.Margins.Top : _options.Margins.Bottom;
        if (marginHeight <= 0.01D) {
            x = y = width = height = 0D;
            return false;
        }

        x = _options.Margins.Left + column * columnWidth;
        width = Math.Max(1D, columnWidth);
        height = Math.Max(0.01D, Math.Min(lineHeight, marginHeight));
        y = top
            ? Math.Max(0D, (marginHeight - height) / 2D)
            : page.Height - marginHeight + Math.Max(0D, (marginHeight - height) / 2D);
        return true;
    }

    private bool TryGetCornerBounds(HtmlRenderPage page, HtmlCssPageMarginPosition position, double lineHeight, out double x, out double y, out double width, out double height) {
        bool left = position == HtmlCssPageMarginPosition.TopLeftCorner || position == HtmlCssPageMarginPosition.BottomLeftCorner;
        bool top = position == HtmlCssPageMarginPosition.TopLeftCorner || position == HtmlCssPageMarginPosition.TopRightCorner;
        double marginWidth = left ? _options.Margins.Left : _options.Margins.Right;
        double marginHeight = top ? _options.Margins.Top : _options.Margins.Bottom;
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
        double marginWidth = left ? _options.Margins.Left : _options.Margins.Right;
        double contentHeight = Math.Max(1D, page.Height - _options.Margins.Top - _options.Margins.Bottom);
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
        y = _options.Margins.Top + section * sectionHeight + Math.Max(0D, (sectionHeight - height) / 2D);
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
