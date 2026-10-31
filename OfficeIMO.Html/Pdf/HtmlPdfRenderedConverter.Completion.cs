using System;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Html.Pdf;

internal static partial class HtmlPdfRenderedConverter {
    // Keep serialization and publication in the same deadline as layout and resource resolution.
    internal static T Complete<T>(HtmlConversionDocument document, HtmlToPdfOptions options,
        CancellationToken cancellationToken, Func<HtmlPdfRenderResult, CancellationToken, T> complete) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        cancellationToken.ThrowIfCancellationRequested();
        HtmlRenderRequest request = HtmlRenderRequest.FromLegacy(
            options, HtmlRenderEncoder.Pdf, HtmlRenderPageSet.All(), forcePrintPaged: true);
        PreparedPdfRequest prepared = PrepareRequest(request, options);
        HtmlRenderOptions resolved = HtmlRenderEngine.PrepareOptions(document, prepared.Request);
        return HtmlRenderEngine.ExecuteWithDeadline(resolved, cancellationToken, token =>
            complete(ConvertCore(document, prepared, resolved, token), token));
    }

    internal static async Task<T> CompleteAsync<T>(HtmlConversionDocument document, HtmlToPdfOptions options,
        CancellationToken cancellationToken, Func<HtmlPdfRenderResult, CancellationToken, Task<T>> complete) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        cancellationToken.ThrowIfCancellationRequested();
        HtmlRenderRequest request = HtmlRenderRequest.FromLegacy(
            options, HtmlRenderEncoder.Pdf, HtmlRenderPageSet.All(), forcePrintPaged: true);
        PreparedPdfRequest prepared = PrepareRequest(request, options);
        HtmlRenderOptions resolved = HtmlRenderEngine.PrepareOptions(document, prepared.Request);
        return await HtmlRenderEngine.ExecuteWithDeadlineAsync(resolved, cancellationToken, async token => {
            HtmlPdfRenderResult rendered = await ConvertCoreAsync(document, prepared, resolved, token).ConfigureAwait(false);
            return await complete(rendered, token).ConfigureAwait(false);
        }).ConfigureAwait(false);
    }
}
