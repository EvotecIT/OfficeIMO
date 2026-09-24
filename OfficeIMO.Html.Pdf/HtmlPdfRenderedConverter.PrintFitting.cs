using OfficeIMO.Drawing;

namespace OfficeIMO.Html.Pdf;

internal static partial class HtmlPdfRenderedConverter {
    private static void ApplyPrintLayoutWidth(HtmlRenderRequest request, HtmlToPdfOptions renderOptions) {
        if (renderOptions.PrintLayoutWidthCssPixels is not double layoutWidth) return;
        if (request.CssMedia != HtmlCssMediaContext.Print
            || request.Surface != HtmlRenderLayoutSurface.Paged
            || request.Pagination != HtmlRenderPaginationPolicy.FragmentedReflow)
            throw new ArgumentException("Print layout fitting requires a paged print reflow request.", nameof(request));
        if (renderOptions.HonorCssPageRules)
            throw new ArgumentException("Print layout fitting requires HonorCssPageRules=false so authored page geometry is not changed.", nameof(renderOptions));
        double physicalWidth = renderOptions.PageWidth;
        if (layoutWidth <= physicalWidth || layoutWidth > renderOptions.MaxSurfaceWidth)
            throw new ArgumentOutOfRangeException(nameof(HtmlToPdfOptions.PrintLayoutWidthCssPixels),
                "The print layout width must exceed the physical page width and stay within the configured surface limit.");
        double layoutScale = physicalWidth / layoutWidth;
        double layoutHeight = renderOptions.PageHeight / layoutScale;
        if (layoutHeight > renderOptions.MaxSurfaceHeight)
            throw new ArgumentOutOfRangeException(nameof(HtmlToPdfOptions.PrintLayoutWidthCssPixels),
                "The scaled page height exceeds the configured surface limit.");

        renderOptions.PrintOutputPageSize = renderOptions.PageSize;
        renderOptions.PageSize = new OfficePageSize(layoutWidth / HtmlRenderOptions.CssPixelsPerInch,
            layoutHeight / HtmlRenderOptions.CssPixelsPerInch);
        HtmlRenderMargins margins = renderOptions.Margins;
        renderOptions.Margins = new HtmlRenderMargins(
            margins.Left / layoutScale,
            margins.Top / layoutScale,
            margins.Right / layoutScale,
            margins.Bottom / layoutScale);
    }

    private static double ResolvePrintLayoutScale(HtmlToPdfOptions options) =>
        options.PrintLayoutWidthCssPixels is double layoutWidth
            ? (options.PrintOutputPageSize ?? options.PageSize).WidthInches * HtmlRenderOptions.CssPixelsPerInch / layoutWidth
            : 1D;
}
