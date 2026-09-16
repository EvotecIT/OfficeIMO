using System.Net;
using AngleSharp;
using AngleSharp.Html.Dom;
using AngleSharp.Io;
using OfficeIMO.Html;

namespace OfficeIMO.Html.Runtime.Worker;

// Root modules join the same source cache as modulepreload and descendant imports.
// Ordinary document resources continue through AngleSharp's requester boundary.
internal sealed class RuntimeDocumentResourceLoader(IBrowsingContext context, RuntimeModuleSourceCache sources,
    Func<RuntimeImportMap?> importMap, HtmlScriptRequest runtimeOptions) : DefaultResourceLoader(context) {
    public override IDownload FetchAsync(ResourceRequest request) {
        if (request.Source is AngleSharp.Dom.IElement element && element.Owner is { } document) {
            SelectResponsiveImage(request, element, document);
            string previousBase=document.BaseUri;
            string currentBase=RuntimeDocumentUrls.Base(document);
            string? attribute=element.LocalName switch { "script" or "img" or "input" or "iframe" or "audio" or "video" or "source"=>"src", "link"=>"href", _=>null };
            string? value=attribute==null ? null : element.GetAttribute(attribute);
            // Rebase direct element resources only. A stylesheet's imported URL or
            // a selected srcset candidate must retain its own request identity.
            if(previousBase!=currentBase && value!=null && new AngleSharp.Dom.Url(new AngleSharp.Dom.Url(previousBase),value).Href==request.Target.Href) {
                var target=new AngleSharp.Dom.Url(new AngleSharp.Dom.Url(currentBase),value);
                request.Target.Href=target.Href;request.Target.Fragment=target.Fragment;request.Target.Query=target.Query;
            }
        }
        if (request.Source is not IHtmlScriptElement script || !string.Equals(script.Type, "module", StringComparison.OrdinalIgnoreCase))
            return base.FetchAsync(request);
        if (script.GetAttribute("crossorigin")?.Equals("use-credentials", StringComparison.OrdinalIgnoreCase) == true)
            throw new HtmlScriptRuntimeException("Credentialed module loading is not supported.");
        var cancellation = new CancellationTokenSource();
        var task = LoadModuleAsync(request.Target, request.IntegritySnapshot, cancellation.Token);
        return new RuntimeModuleDownload(new AngleSharp.Dom.Url(request.Target.Href), request.Source, task, cancellation);
    }

    private static void SetTarget(AngleSharp.Dom.Url target, string address) {
        var selected = new AngleSharp.Dom.Url(address);
        target.Href = selected.Href;
        target.Fragment = selected.Fragment;
        target.Query = selected.Query;
    }

    private void SelectResponsiveImage(ResourceRequest request, AngleSharp.Dom.IElement element,
        AngleSharp.Dom.IDocument document) {
        if (element is not IHtmlImageElement image ||
            !string.Equals(image.ParentElement?.LocalName, "picture", StringComparison.OrdinalIgnoreCase)) return;
        var renderOptions = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = runtimeOptions.ViewportWidth,
            ViewportHeight = runtimeOptions.ViewportHeight
        };
        string? selected = HtmlImageSourceResolver.ResolveImageSourceCandidatesForRendering(
            image, new Uri(RuntimeDocumentUrls.Base(document)), HtmlUrlPolicy.CreateWebResourceProfile(), renderOptions)
            .FirstOrDefault();
        if (!string.IsNullOrWhiteSpace(selected)) SetTarget(request.Target, selected);
    }

    private async Task<IResponse> LoadModuleAsync(AngleSharp.Dom.Url target, IntegrityMetadataSnapshot? integritySnapshot,
        CancellationToken token) {
        var url = new Uri(target.Href);
        string? integrityMetadata = integritySnapshot != null
            ? integritySnapshot.Resolve(importMap()?.IntegrityFor(url))
            : importMap()?.IntegrityFor(url);
        RuntimeModuleSource source = await sources.GetOrLoad(url.AbsoluteUri, url, integrityMetadata, token).ConfigureAwait(false);
        return new DefaultResponse {
            Address = new AngleSharp.Dom.Url(source.Location),
            StatusCode = (HttpStatusCode)source.StatusCode,
            Content = new MemoryStream(source.Buffer, writable: false),
            Headers = new Dictionary<string, string>(source.Headers, StringComparer.OrdinalIgnoreCase) {
                ["Content-Type"] = source.ContentType
            }
        };
    }
}
