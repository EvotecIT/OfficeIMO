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
    HtmlScriptRequest page = incoming.Page.Snapshot();
    if (page.Profile != HtmlRuntimeProfile.WebApplicationV1 || page.ResourcePolicy.AllowNetwork)
        throw new NotSupportedException("The isolated renderer accepts only offline WebApplicationV1 input.");
    IHtmlRuntimeHost host = new HtmlProcessRuntimeProvider(workerPath, AngleSharpDomServices.Instance);
    var rendering = new HtmlToPdfOptions { ViewportWidth = 816D, Margins = HtmlRenderMargins.All(0D) };
    HtmlApplicationDocumentResult result = await HtmlApplicationDocumentWorkflow.RunAsync(host,
        new HtmlApplicationDocumentRequest {
            Page = page,
            RenderRequests = new[] {
                HtmlRenderRequest.Create(HtmlRenderIntentProfile.ScreenFullPage, HtmlRenderEncoder.Png, rendering),
                HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf, rendering),
                HtmlRenderRequest.Create(HtmlRenderIntentProfile.ScreenSnapshotPaged, HtmlRenderEncoder.Pdf, rendering)
            }
        }, deadline.Token);
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
}
await HtmlRuntimeProtocol.WriteAsync(output, response, 24 * 1024 * 1024, CancellationToken.None);
