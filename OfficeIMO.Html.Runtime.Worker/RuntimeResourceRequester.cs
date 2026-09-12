using AngleSharp;
using AngleSharp.Dom;
using AngleSharp.Io;
using System.Net;

namespace OfficeIMO.Html.Runtime.Worker;

internal sealed class RuntimeResourceRequester(RuntimeResourceLoader loader, RuntimeScriptErrors errors) : BaseRequester {
    public override bool SupportsProtocol(string protocol) => true; // Route every scheme through the owned policy, including refusals.

    protected override async Task<IResponse?> PerformRequestAsync(Request request, CancellationToken cancel) {
        try {
            if (request.Method != AngleSharp.Io.HttpMethod.Get) throw new HtmlScriptRuntimeException("Document resources require GET requests.");
            var url = new Uri(request.Address.Href);
            var resource = await (RuntimeDocumentResourceLoader.IsModule(request)
                ? loader.FetchAsync(url, new RuntimeFetchRequest(), cancel) : loader.LoadAsync(url, cancel)).ConfigureAwait(false);
            if (resource.StatusCode < 200 || resource.StatusCode >= 300) throw new HtmlScriptRuntimeException("Document resource loading returned HTTP " + resource.StatusCode + ".");
            return new DefaultResponse {
                Address = new Url(resource.FinalUrl.AbsoluteUri),
                StatusCode = (HttpStatusCode)resource.StatusCode,
                Content = new MemoryStream(resource.Buffer, writable: false),
                Headers = new Dictionary<string, string>(resource.Headers, StringComparer.OrdinalIgnoreCase) { ["Content-Type"] = resource.ContentType }
            };
        } catch (Exception error) { errors.Report(error.Message); throw; }
    }
}
