namespace OfficeIMO.Html.Runtime;

// Private protocol between the public-page pilot and its whole-pipeline OCI
// process. No untrusted HTML is parsed, scripted or rendered on the host.
internal sealed class HtmlPublicRenderRequest {
    public HtmlScriptRequest Page { get; set; } = new();
}

internal sealed class HtmlPublicResourceBatch {
    public HtmlRuntimeResource[] Resources { get; set; } = Array.Empty<HtmlRuntimeResource>();
}

internal enum HtmlPublicRenderStage {
    ResourceDiscovery,
    Rendering,
    Output
}

internal sealed class HtmlPublicRenderResponse {
    public HtmlPublicRenderStage Stage { get; set; }
    public string RendererSha256 { get; set; } = string.Empty;
    public string WorkerSha256 { get; set; } = string.Empty;
    public string RendererFilesSha256 { get; set; } = string.Empty;
    public string WorkerFilesSha256 { get; set; } = string.Empty;
    public bool DiscoveryComplete { get; set; }
    public string[] DiscoveryUrls { get; set; } = Array.Empty<string>();
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
