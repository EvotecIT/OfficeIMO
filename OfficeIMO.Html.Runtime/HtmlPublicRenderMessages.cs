namespace OfficeIMO.Html.Runtime;

// Private protocol between the public-page pilot and its whole-pipeline OCI
// process. No untrusted HTML is parsed, scripted or rendered on the host.
internal sealed class HtmlPublicRenderRequest {
    public HtmlScriptRequest Page { get; set; } = new();
    public long MaxOutputBytesPerArtifact { get; set; } = 8L * 1024 * 1024;
    public long MaxTotalOutputBytes { get; set; } = 12L * 1024 * 1024;
}

internal sealed class HtmlPublicResourceBatch {
    public HtmlRuntimeResource[] Resources { get; set; } = Array.Empty<HtmlRuntimeResource>();
    public HtmlRuntimeFetchReplay[] FetchReplays { get; set; } = Array.Empty<HtmlRuntimeFetchReplay>();
}

internal static class HtmlRuntimeFetchTranscript {
    internal static void Validate(IReadOnlyList<HtmlRuntimeFetchReplay> supplied,
        IReadOnlyList<string> consumedIdentities) {
        ArgumentNullException.ThrowIfNull(supplied);
        ArgumentNullException.ThrowIfNull(consumedIdentities);
        string[] expected = supplied.Select(replay => replay?.Identity
            ?? throw new HtmlScriptRuntimeException("The dynamic replay transcript contains a null replay.")).ToArray();
        if (!expected.SequenceEqual(consumedIdentities, StringComparer.Ordinal))
            throw new HtmlScriptRuntimeException("The isolated execution did not consume the acquired dynamic replay transcript exactly.");
    }
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
    public HtmlRuntimeFetchDiscovery[] DiscoveryRequests { get; set; } = Array.Empty<HtmlRuntimeFetchDiscovery>();
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
