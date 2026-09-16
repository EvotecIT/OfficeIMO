using OfficeIMO.Html.Pdf;
using OfficeIMO.Html.Runtime;

namespace OfficeIMO.Html.Runtime.Rendering;

/// <summary>Composes a trusted application session with independent OfficeIMO rendering outputs.</summary>
public static class HtmlApplicationDocumentWorkflow {
    /// <summary>Runs ordered actions, captures one ready document, closes the live context, then renders each requested output.</summary>
    public static async Task<HtmlApplicationDocumentResult> RunAsync(IHtmlRuntimeHost host,
        HtmlApplicationDocumentRequest request, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(host);
        HtmlApplicationDocumentRequest input = (request ?? throw new ArgumentNullException(nameof(request))).Snapshot();
        foreach (HtmlRenderRequest render in input.RenderRequests) {
            if (render.Options.ResourceResolver != null) throw new ArgumentException(
                "Render requests in this workflow must use the supplied and captured resource set, not an external resolver.", nameof(request));
        }
        HtmlRuntimeProviderDescriptor descriptor = host.Descriptor;
        if (!descriptor.Profiles.Contains(input.Page.Profile)) throw new NotSupportedException("The selected runtime does not support this page profile.");
        if (!descriptor.Supports(HtmlRuntimeCapabilityIds.OperationTrace)) throw new NotSupportedException("The selected runtime does not expose operation traces.");
        if (input.Actions.Count > 0 && !descriptor.Supports(HtmlRuntimeCapabilityIds.StructuredActions))
            throw new NotSupportedException("The selected runtime does not support structured actions.");

        HtmlScriptCapture capture;
        HtmlRuntimeTrace trace;
        HtmlRuntimeProviderDescriptor provider;
        var actionResults = new List<HtmlAutomationResult>(input.Actions.Count);
        await using (IHtmlRuntimeContext context = await host.CreateContextAsync(input.Context, cancellationToken).ConfigureAwait(false)) {
            await using (IHtmlRuntimePage page = await context.OpenPageAsync(input.Page, cancellationToken).ConfigureAwait(false)) {
                provider = page.Provider;
                foreach (HtmlAutomationRequest action in input.Actions) {
                    actionResults.Add((await page.AutomateAsync(action, cancellationToken).ConfigureAwait(false)).EnsureSuccess());
                }
                capture = await page.CaptureAsync(input.FinalReadyExpression, cancellationToken).ConfigureAwait(false);
                trace = page.GetTrace();
            }
        }

        HtmlConversionDocument document = HtmlConversionDocument.FromDocument(capture.CreateStandaloneDocument(),
            new HtmlConversionDocumentOptions { BaseUri = capture.BaseUri });
        IReadOnlyDictionary<string, HtmlRuntimeResource> resources = RetainRenderResources(input.Page, capture.Resources);
        var outputs = new List<HtmlApplicationRenderOutput>(input.RenderRequests.Count);
        foreach (HtmlRenderRequest requested in input.RenderRequests) {
            cancellationToken.ThrowIfCancellationRequested();
            HtmlRenderOptions options = requested.Options;
            if (requested.Encoder == HtmlRenderEncoder.Pdf && options is not HtmlToPdfOptions) {
                options = new HtmlToPdfOptions(options);
            }
            options.BaseUri = capture.BaseUri;
            options.ResourceResolver = (resourceRequest, _) => Task.FromResult(resources.TryGetValue(
                resourceRequest.Uri.AbsoluteUri, out HtmlRuntimeResource? resource)
                ? new HtmlResolvedResource(resource.Content, resource.ContentType, resource.FinalUrl, resource.RedirectCount)
                : null);
            if (options is HtmlToPdfOptions pdfOptions) {
                pdfOptions.ResourcePolicy.AllowRemoteResourceResolution = true;
            }
            HtmlRenderRequest render = requested.WithOptions(options).WithDocumentState(HtmlRenderDocumentState.RuntimeSnapshot);
            if (render.Encoder == HtmlRenderEncoder.Pdf) {
                HtmlPdfRenderRequestResult pdf = await document.RenderToPdfResultAsync(render, cancellationToken).ConfigureAwait(false);
                outputs.Add(new HtmlApplicationRenderOutput(pdf.RenderResult, pdf: pdf));
                continue;
            }
            HtmlRenderResult result = await HtmlRenderEngine.ExecuteAsync(document, render, cancellationToken).ConfigureAwait(false);
            if (render.Encoder is HtmlRenderEncoder.Png or HtmlRenderEncoder.Jpeg or HtmlRenderEncoder.Tiff or HtmlRenderEncoder.Webp or HtmlRenderEncoder.Svg) {
                outputs.Add(new HtmlApplicationRenderOutput(result, result.ExportImages(cancellationToken)));
            } else {
                outputs.Add(new HtmlApplicationRenderOutput(result));
            }
        }
        return new HtmlApplicationDocumentResult(provider, capture, trace,
            Array.AsReadOnly(resources.Values.Distinct().ToArray()),
            Array.AsReadOnly(actionResults.ToArray()), Array.AsReadOnly(outputs.ToArray()));
    }

    private static IReadOnlyDictionary<string, HtmlRuntimeResource> RetainRenderResources(
        HtmlScriptRequest page, IReadOnlyList<HtmlRuntimeResource> captured) {
        HtmlRuntimeResourcePolicy policy = page.ResourcePolicy;
        var origins = new HashSet<string>(policy.AllowedOrigins.Select(origin => origin.GetLeftPart(UriPartial.Authority)),
            StringComparer.OrdinalIgnoreCase) { page.DocumentUrl.GetLeftPart(UriPartial.Authority) };
        var resources = new Dictionary<string, HtmlRuntimeResource>(StringComparer.Ordinal);
        bool Allowed(HtmlRuntimeResource resource) =>
            resource.StatusCode is >= 200 and < 300 && resource.Length > 0 &&
            resource.RedirectCount <= policy.MaxRedirects &&
            origins.Contains(resource.Url.GetLeftPart(UriPartial.Authority)) &&
            origins.Contains(resource.FinalUrl.GetLeftPart(UriPartial.Authority));

        // Direct URL identities always outrank redirect aliases, regardless of input order.
        foreach (HtmlRuntimeResource resource in page.Resources) {
            if (!Allowed(resource)) continue;
            resources[resource.Url.AbsoluteUri] = resource;
        }
        foreach (HtmlRuntimeResource resource in captured) {
            // The observed response wins, including a failure that invalidates supplied bytes.
            resources.Remove(resource.Url.AbsoluteUri);
            if (!Allowed(resource)) continue;
            resources[resource.Url.AbsoluteUri] = resource;
        }
        foreach (HtmlRuntimeResource resource in captured.Concat(page.Resources)) {
            if (Allowed(resource) && resources.TryGetValue(resource.Url.AbsoluteUri, out HtmlRuntimeResource? direct)
                && ReferenceEquals(direct, resource)) {
                resources.TryAdd(resource.FinalUrl.AbsoluteUri, resource);
            }
        }
        return resources;
    }
}
