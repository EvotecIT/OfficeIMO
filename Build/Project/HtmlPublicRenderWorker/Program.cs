using System.Security.Cryptography;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime;
using OfficeIMO.Html.Runtime.Rendering;

using Stream input = Console.OpenStandardInput();
using Stream output = Console.OpenStandardOutput();
string rendererPath = typeof(Program).Assembly.Location;
string workerPath = Path.Combine(AppContext.BaseDirectory, "worker", "OfficeIMO.Html.Runtime.Worker.dll");
var response = new HtmlPublicRenderResponse {
    RendererSha256 = Convert.ToHexStringLower(SHA256.HashData(File.ReadAllBytes(rendererPath))),
    WorkerSha256 = Convert.ToHexStringLower(SHA256.HashData(File.ReadAllBytes(workerPath))),
    RendererFilesSha256 = HtmlPublicArtifactDigest.DirectorySha256(AppContext.BaseDirectory, "worker"),
    WorkerFilesSha256 = HtmlPublicArtifactDigest.DirectorySha256(Path.GetDirectoryName(workerPath)!)
};
try {
    using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(90));
    HtmlPublicRenderRequest incoming = await HtmlRuntimeProtocol.ReadAsync<HtmlPublicRenderRequest>(
        input, 24 * 1024 * 1024, deadline.Token)
        ?? throw new HtmlScriptRuntimeException("The isolated render request is missing.");
    var rendering = new HtmlToPdfOptions { ViewportWidth = 816D, ViewportHeight = 720D,
        Margins = HtmlRenderMargins.All(0D) };
    incoming.Page.ViewportWidth = rendering.ViewportWidth;
    incoming.Page.ViewportHeight = rendering.ViewportHeight ?? 720D;
    rendering.MediaFeatures.ResolutionDpi = incoming.Page.DevicePixelRatio * HtmlRenderOptions.CssPixelsPerInch;
    HtmlScriptRequest page = incoming.Page.Snapshot();
    if (page.Profile != HtmlRuntimeProfile.WebApplicationV1 || page.ResourcePolicy.AllowNetwork)
        throw new NotSupportedException("The isolated renderer accepts only offline WebApplicationV1 input.");
    var discovery = new HtmlApplicationResourceDiscovery(page.ViewportWidth, page.ViewportHeight, page.DevicePixelRatio);
    var supplied = page.Resources.ToList();
    string[] pending = discovery.DiscoverDocument(page.Html, page.DocumentUrl, supplied);
    IHtmlRuntimeHost host = new HtmlProcessRuntimeProvider(workerPath, AngleSharpDomServices.Instance);
    HtmlApplicationDocumentResult result;
    var missingAtRuntime = new HashSet<string>(StringComparer.Ordinal);
    for (int round = 0; ; ) {
        if (pending.Length != 0) {
            if (++round > 16)
                throw new HtmlScriptRuntimeException("Resource discovery exceeded its round limit.");
            response.DiscoveryUrls = pending;
            response.DiscoveryComplete = false;
            await HtmlRuntimeProtocol.WriteAsync(output, response, 24 * 1024 * 1024, deadline.Token);
            HtmlPublicResourceBatch batch = await HtmlRuntimeProtocol.ReadAsync<HtmlPublicResourceBatch>(
                input, 24 * 1024 * 1024, deadline.Token)
                ?? throw new HtmlScriptRuntimeException("The host did not answer resource discovery.");
            if (batch.Resources == null || batch.Resources.Length > 24 - supplied.Count)
                throw new HtmlScriptRuntimeException("The host exceeded the pilot's supplied resource limit.");
            var requested = pending.ToHashSet(StringComparer.Ordinal);
            foreach (HtmlRuntimeResource resource in batch.Resources) {
                if (resource == null || !requested.Remove(HtmlRuntimeResourcePolicy.Key(resource.Url)))
                    throw new HtmlScriptRuntimeException("The host supplied an unexpected or duplicate discovered resource.");
                supplied.Add(resource);
            }
            page.Resources = supplied.ToArray();
            page.ResourcePolicy.AllowedOrigins = page.ResourcePolicy.AllowedOrigins
                .Concat(supplied.SelectMany(resource => new[] { resource.Url, resource.FinalUrl }))
                .Select(resourceUrl => new Uri(resourceUrl.GetLeftPart(UriPartial.Authority)))
                .Where(origin => origin.GetLeftPart(UriPartial.Authority) != page.DocumentUrl.GetLeftPart(UriPartial.Authority))
                .DistinctBy(origin => origin.AbsoluteUri, StringComparer.OrdinalIgnoreCase).ToArray();
            page = page.Snapshot();
            pending = discovery.DiscoverResources(batch.Resources);
            continue;
        }
        try {
            result = await HtmlApplicationDocumentWorkflow.RunAsync(host,
                new HtmlApplicationDocumentRequest {
                    Page = page,
                    Context = new HtmlRuntimeContextOptions {
                        Trace = new HtmlRuntimeTraceOptions { IncludeUrls = true }
                    },
                    RenderRequests = new[] {
                        HtmlRenderRequest.Create(HtmlRenderIntentProfile.ScreenFullPage, HtmlRenderEncoder.Png, rendering),
                        HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf, rendering),
                        HtmlRenderRequest.Create(HtmlRenderIntentProfile.ScreenSnapshotPaged, HtmlRenderEncoder.Pdf, rendering)
                    }
                }, deadline.Token);
            if (result.Trace.IsTruncated)
                throw new HtmlScriptRuntimeException("Resource discovery trace exceeded its event limit.");
            pending = result.Trace.Events
                .Where(entry => entry.Kind == HtmlRuntimeEventKind.Policy && entry.Operation == "network-access" &&
                    entry.Status == "blocked" && entry.Decision == "network-disabled-replayable-get" &&
                    entry.Method == "GET" && entry.Url != null)
                .Select(entry => HtmlRuntimeResourcePolicy.Key(entry.Url!))
                .Where(missingAtRuntime.Add).ToArray();
            if (pending.Length != 0) continue;
            break;
        } catch (HtmlScriptRuntimeException error) when (error.MissingResourceUrls.Count != 0) {
            pending = error.MissingResourceUrls.Select(HtmlRuntimeResourcePolicy.Key)
                .Where(missingAtRuntime.Add).ToArray();
            if (pending.Length == 0) throw;
        }
    }
    response.DiscoveryUrls = Array.Empty<string>();
    response.DiscoveryComplete = true;
    await HtmlRuntimeProtocol.WriteAsync(output, response, 24 * 1024 * 1024, deadline.Token);
    byte[] screen = result.Outputs[0].Images.Single().Bytes;
    byte[] print = result.Outputs[1].Pdf!.ToBytes();
    byte[] screenToPage = result.Outputs[2].Pdf!.ToBytes();
    if (screen.Length > 8 * 1024 * 1024 || print.Length > 8 * 1024 * 1024 ||
        screenToPage.Length > 8 * 1024 * 1024 ||
        (long)screen.Length + print.Length + screenToPage.Length > 12 * 1024 * 1024)
        throw new HtmlScriptRuntimeException("The isolated render output exceeds its byte budget.");
    response.ProviderId = result.Provider.Id;
    response.CaptureUrl = result.Capture.DocumentUrl.AbsoluteUri;
    response.CaptureManifest = result.Capture.ArtifactManifest.Id;
    response.TraceEntries = result.Trace.Events.Take(128).Select(entry =>
        entry.Kind + ":" + entry.Operation + ":" + entry.Status +
        (entry.Detail == null ? string.Empty : ":" + entry.Detail)).ToArray();
    response.Screen = screen;
    response.Print = print;
    response.ScreenToPage = screenToPage;
} catch (Exception error) {
    response.ErrorKind = error.GetType().Name;
    response.Error = error.Message.Length > 1024 ? error.Message[..1024] : error.Message;
    response.DiscoveryComplete = true;
    response.DiscoveryUrls = Array.Empty<string>();
}
await HtmlRuntimeProtocol.WriteAsync(output, response, 24 * 1024 * 1024, CancellationToken.None);
