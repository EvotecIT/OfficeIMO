using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private IReadOnlyList<HtmlRenderPage> ApplyPageMarginContent(IReadOnlyList<HtmlRenderPage> pages) {
        var rendered = new List<HtmlRenderPage>(pages.Count);
        foreach (HtmlRenderPage page in pages) {
            IReadOnlyDictionary<HtmlCssPageMarginPosition, HtmlCssPageMarginTemplate> boxes = _pageRules.ResolveMarginBoxes(page.PageNumber);
            if (boxes.Count == 0) {
                rendered.Add(page);
                continue;
            }

            var visuals = new List<HtmlRenderVisual>(page.Visuals);
            foreach (HtmlCssPageMarginTemplate box in boxes.Values.OrderBy(item => item.Position)) {
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

            rendered.Add(new HtmlRenderPage(page.PageNumber, page.Width, page.Height, visuals));
        }

        return rendered.AsReadOnly();
    }

    private bool TryGetMarginBoxBounds(HtmlRenderPage page, HtmlCssPageMarginPosition position, double fontSize, out double x, out double y, out double width, out double height) {
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

        double lineHeight = Math.Max(1D, fontSize * _options.DefaultLineHeight);
        x = _options.Margins.Left + column * columnWidth;
        width = Math.Max(1D, columnWidth);
        height = Math.Max(0.01D, Math.Min(lineHeight, marginHeight));
        y = top
            ? Math.Max(0D, (marginHeight - height) / 2D)
            : page.Height - marginHeight + Math.Max(0D, (marginHeight - height) / 2D);
        return true;
    }

    private static string GetMarginBoxName(HtmlCssPageMarginPosition position) {
        switch (position) {
            case HtmlCssPageMarginPosition.TopLeft: return "top-left";
            case HtmlCssPageMarginPosition.TopCenter: return "top-center";
            case HtmlCssPageMarginPosition.TopRight: return "top-right";
            case HtmlCssPageMarginPosition.BottomLeft: return "bottom-left";
            case HtmlCssPageMarginPosition.BottomCenter: return "bottom-center";
            default: return "bottom-right";
        }
    }
}
