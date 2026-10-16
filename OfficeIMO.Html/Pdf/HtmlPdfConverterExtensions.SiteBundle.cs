using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Html.Pdf;

/// <summary>Renders bounded ZIP site bundles through the existing HTML/PDF request and resource owners.</summary>
public static class HtmlSiteBundlePdfExtensions {
    /// <summary>Renders a PDF intent and retains HTML surface mapping and the combined PDF report.
    /// Embedded resources require AllowEmbeddedPackageResources; host reads remain separately controlled.</summary>
    public static HtmlPdfRenderRequestResult RenderToPdfResult(this HtmlSiteBundle bundle,
        HtmlRenderRequest request, CancellationToken cancellationToken = default) {
        if (bundle == null) throw new ArgumentNullException(nameof(bundle));
        if (request == null) throw new ArgumentNullException(nameof(request));
        return bundle.HtmlDocument.RenderToPdfResult(Prepare(bundle, request), cancellationToken);
    }

    /// <summary>Asynchronously renders a PDF intent with the archive's snapshot resources, without
    /// implicitly enabling local or remote resource access.</summary>
    public static Task<HtmlPdfRenderRequestResult> RenderToPdfResultAsync(this HtmlSiteBundle bundle,
        HtmlRenderRequest request, CancellationToken cancellationToken = default) {
        if (bundle == null) throw new ArgumentNullException(nameof(bundle));
        if (request == null) throw new ArgumentNullException(nameof(request));
        return bundle.HtmlDocument.RenderToPdfResultAsync(Prepare(bundle, request), cancellationToken);
    }

    private static HtmlRenderRequest Prepare(HtmlSiteBundle bundle, HtmlRenderRequest request) {
        var options = new HtmlToPdfOptions(request.Options);
        options.BaseUri ??= bundle.BaseUri;
        options.EmbeddedPackageHostResourceUrlPolicy = options.GetResourceUrlPolicy().Clone();
        options.EmbeddedPackageResourceResolver = options.ResourcePolicy.AllowEmbeddedPackageResources
            ? bundle.CreateResourceResolver()
            : null;
        // Sync rendering uses the same immutable archive bytes. Do not carry a host
        // synchronous resolver through the embedded-resource policy boundary.
        options.SynchronousResourceResolver = options.ResourcePolicy.AllowEmbeddedPackageResources
            ? bundle.TryResolveResource
            : null;
        return request.WithOptions(options);
    }
}
