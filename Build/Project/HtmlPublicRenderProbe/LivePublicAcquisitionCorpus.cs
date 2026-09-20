using System.Net;
using System.Text;
using OfficeIMO.Html.Runtime;

internal sealed record LivePublicAcquisitionCorpus(IReadOnlyList<LivePublicAcquisitionResult> Results) {
    internal static async Task<LivePublicAcquisitionCorpus> RunAsync() {
        var results = new List<LivePublicAcquisitionResult> {
            await RedirectAsync(),
            await CorsPreflightAsync()
        };
        return new LivePublicAcquisitionCorpus(results.AsReadOnly());
    }

    private static async Task<LivePublicAcquisitionResult> RedirectAsync() {
        const string name = "live-public-same-host-redirect";
        Uri requested = new("https://httpbingo.org/redirect/1");
        HtmlPublicResourceResult? acquired = null;
        Exception? failure = null;
        try {
            var broker = new HtmlPublicResourceBroker(new[] { requested.IdnHost });
            acquired = await broker.FetchAsync(requested);
            if (acquired.Redirects.Count != 1 || acquired.Resource.FinalUrl != new Uri("https://httpbingo.org/get")
                || acquired.Redirects[0].StatusCode != 302)
                throw new IOException("The live public redirect did not retain the expected single validated hop.");
        } catch (Exception error) { failure = error; }
        return Result(name, requested, acquired, failure);
    }

    private static async Task<LivePublicAcquisitionResult> CorsPreflightAsync() {
        const string name = "live-public-cors-preflight";
        Uri requested = new("https://httpbingo.org/anything/officeimo-live-preflight");
        HtmlPublicResourceResult? acquired = null;
        Exception? failure = null;
        try {
            var broker = new HtmlPublicResourceBroker(new[] { "example.com" },
                dynamicOrigins: new[] { new Uri("https://httpbingo.org/") });
            var request = new HtmlRuntimeFetchRequest(requested, new Uri("https://example.com/"), "POST",
                new Dictionary<string, string> { ["Content-Type"] = "application/json" },
                Encoding.UTF8.GetBytes("{}"), credentials: "omit");
            acquired = await broker.FetchAsync(new HtmlRuntimeFetchDiscovery(request, 1));
            if (acquired.HttpExchanges?.Select(exchange => exchange.Method)
                    .SequenceEqual(new[] { "OPTIONS", "POST" }, StringComparer.Ordinal) != true
                || acquired.DynamicHops?.Single().PreflightResponse?.StatusCode is not 200)
                throw new IOException("The live public endpoint did not complete the expected OPTIONS and POST exchange.");
        } catch (Exception error) { failure = error; }
        return Result(name, requested, acquired, failure);
    }

    private static LivePublicAcquisitionResult Result(string name, Uri requested,
        HtmlPublicResourceResult? acquired, Exception? failure) => new(name, failure == null,
            requested.AbsoluteUri, acquired?.Resource.FinalUrl.AbsoluteUri,
            acquired?.ConnectedAddress.ToString(), acquired?.Resource.StatusCode,
            acquired?.Sha256, acquired?.Redirects.Select(redirect => new AcquisitionRedirect(
                redirect.From.AbsoluteUri, redirect.To.AbsoluteUri, redirect.StatusCode,
                redirect.ConnectedAddress.ToString())).ToArray() ?? [],
            acquired?.HttpExchanges?.Select(exchange => new LivePublicHttpExchange(
                exchange.Url.AbsoluteUri, exchange.Method, exchange.StatusCode,
                exchange.RequestBodyByteCount, exchange.RequestBodySha256,
                exchange.ResponseByteCount, exchange.ResponseSha256, exchange.ConnectedAddress.ToString())).ToArray() ?? [],
            failure?.GetType().Name, failure?.Message);
}

internal sealed record LivePublicAcquisitionResult(string Name, bool Passed, string RequestedUrl,
    string? FinalUrl, string? ConnectedAddress, int? StatusCode, string? Sha256,
    AcquisitionRedirect[] Redirects, LivePublicHttpExchange[] Exchanges, string? ErrorKind, string? Error);

internal sealed record LivePublicHttpExchange(string Url, string Method, int StatusCode,
    long RequestBodyByteCount, string? RequestBodySha256, long ResponseByteCount,
    string ResponseSha256, string ConnectedAddress);
