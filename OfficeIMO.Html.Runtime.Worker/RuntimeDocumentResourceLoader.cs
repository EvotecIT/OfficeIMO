using System.Runtime.CompilerServices;
using AngleSharp;
using AngleSharp.Html.Dom;
using AngleSharp.Io;

namespace OfficeIMO.Html.Runtime.Worker;

// Carry module request intent to the shared requester without putting an internal
// marker on the network or conflating script requests with ordinary images/styles.
internal sealed class RuntimeDocumentResourceLoader(IBrowsingContext context) : DefaultResourceLoader(context) {
    private static readonly ConditionalWeakTable<Request, object> ModuleRequests = new();
    private static readonly object Marker = new();
    internal static bool IsModule(Request request) => ModuleRequests.TryGetValue(request, out _);

    public override IDownload FetchAsync(ResourceRequest request) {
        if (request.Source is AngleSharp.Dom.IElement element && element.Owner is { } document) {
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
        var data = new Request { Address = request.Target, Content = Stream.Null, Method = AngleSharp.Io.HttpMethod.Get };
        ModuleRequests.Add(data, Marker);
        return DownloadAsync(data, request.Source);
    }
}
