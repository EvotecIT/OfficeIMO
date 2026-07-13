using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

public static partial class HtmlImageExportExtensions {
    /// <summary>Renders one selected surface to the requested image format with dimensions and diagnostics.</summary>
    public static OfficeImageExportResult ExportImage(this HtmlConversionDocument document, OfficeImageExportFormat format, HtmlRenderOptions? options = null, int pageIndex = 0) {
        HtmlRenderOptions resolved = Normalize(options, pageIndex);
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(document, resolved);
        if (pageIndex >= rendered.Pages.Count) throw new ArgumentOutOfRangeException(nameof(pageIndex), "The selected HTML render page does not exist.");
        return RenderPage(rendered.Pages[pageIndex], format, resolved, rendered.DiagnosticReport, CancellationToken.None);
    }

    /// <summary>Renders all surfaces to the requested image format.</summary>
    public static IReadOnlyList<OfficeImageExportResult> ExportImages(this HtmlConversionDocument document, OfficeImageExportFormat format, HtmlRenderOptions? options = null) {
        HtmlRenderOptions resolved = Normalize(options, 0);
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(document, resolved);
        var results = new List<OfficeImageExportResult>(rendered.Pages.Count);
        foreach (HtmlRenderPage page in rendered.Pages) results.Add(RenderPage(page, format, resolved, rendered.DiagnosticReport, CancellationToken.None));
        return results.AsReadOnly();
    }

    /// <summary>Asynchronously renders one selected surface to the requested image format.</summary>
    public static async Task<OfficeImageExportResult> ExportImageAsync(this HtmlConversionDocument document, OfficeImageExportFormat format, HtmlRenderOptions? options = null, int pageIndex = 0, CancellationToken cancellationToken = default) {
        HtmlRenderOptions resolved = Normalize(options, pageIndex);
        HtmlRenderDocument rendered = await HtmlRenderEngine.RenderAsync(document, resolved, cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        if (pageIndex >= rendered.Pages.Count) throw new ArgumentOutOfRangeException(nameof(pageIndex), "The selected HTML render page does not exist.");
        return RenderPage(rendered.Pages[pageIndex], format, resolved, rendered.DiagnosticReport, cancellationToken);
    }

    /// <summary>Asynchronously renders all surfaces to the requested image format.</summary>
    public static async Task<IReadOnlyList<OfficeImageExportResult>> ExportImagesAsync(this HtmlConversionDocument document, OfficeImageExportFormat format, HtmlRenderOptions? options = null, CancellationToken cancellationToken = default) {
        HtmlRenderOptions resolved = Normalize(options, 0);
        HtmlRenderDocument rendered = await HtmlRenderEngine.RenderAsync(document, resolved, cancellationToken).ConfigureAwait(false);
        var results = new List<OfficeImageExportResult>(rendered.Pages.Count);
        foreach (HtmlRenderPage page in rendered.Pages) {
            cancellationToken.ThrowIfCancellationRequested();
            results.Add(RenderPage(page, format, resolved, rendered.DiagnosticReport, cancellationToken));
        }
        return results.AsReadOnly();
    }

}
