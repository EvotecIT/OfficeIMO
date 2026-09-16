namespace OfficeIMO.Html.Runtime;

// Private protocol between the public-page pilot and its whole-pipeline OCI
// process. No untrusted HTML is parsed, scripted or rendered on the host.
internal sealed class HtmlPublicRenderRequest {
    public HtmlScriptRequest Page { get; set; } = new();
}

internal sealed class HtmlPublicRenderResponse {
    public string RendererSha256 { get; set; } = string.Empty;
    public string WorkerSha256 { get; set; } = string.Empty;
    public string RendererFilesSha256 { get; set; } = string.Empty;
    public string WorkerFilesSha256 { get; set; } = string.Empty;
    public string? ErrorKind { get; set; }
    public string? Error { get; set; }
    public string? ProviderId { get; set; }
    public string? CaptureUrl { get; set; }
    public string? CaptureManifest { get; set; }
    public string[] TraceEntries { get; set; } = Array.Empty<string>();
    public byte[]? Screen { get; set; }
    public byte[]? Print { get; set; }
    public byte[]? ScreenToPage { get; set; }
}
