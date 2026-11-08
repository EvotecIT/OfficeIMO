using System.Text;
using OfficeIMO.Html;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html.Pdf;

public static partial class HtmlImageExportExtensions {
    /// <summary>Renders the selected page as decoded SVG text.</summary>
    public static string ToSvg(this HtmlConversionDocument document, HtmlRenderOptions? options = null,
        CancellationToken cancellationToken = default) => document.ToSvg(options, out _, cancellationToken);

    /// <summary>Renders SVG text and retains the selected page's fidelity report.</summary>
    public static string ToSvg(this HtmlConversionDocument document, HtmlRenderOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        HtmlRenderOptions resolved = options?.Clone() ?? new HtmlRenderOptions();
        resolved.Format = OfficeImageExportFormat.Svg;
        return Encoding.UTF8.GetString(document.ToImage(resolved, out report, cancellationToken).Bytes);
    }

    /// <summary>Asynchronously renders SVG text with resource resolution and its completed report.</summary>
    public static async Task<OfficeConversionResult<string>> ToSvgAsync(this HtmlConversionDocument document,
        HtmlRenderOptions? options = null, CancellationToken cancellationToken = default) {
        HtmlRenderOptions resolved = options?.Clone() ?? new HtmlRenderOptions();
        resolved.Format = OfficeImageExportFormat.Svg;
        OfficeImageExportResult result = await document.ToImageAsync(resolved, cancellationToken).ConfigureAwait(false);
        return new OfficeConversionResult<string>(Encoding.UTF8.GetString(result.Bytes), result.CreateReport());
    }
}
