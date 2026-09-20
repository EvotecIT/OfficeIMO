namespace OfficeIMO.Html.Runtime;

// Private protocol between the public-page pilot and its whole-pipeline OCI
// process. No untrusted HTML is parsed, scripted or rendered on the host.
internal sealed class HtmlPublicRenderRequest {
    internal const int MaximumProtocolCharacters = 24 * 1024 * 1024;
    public HtmlScriptRequest Page { get; set; } = new();
    public HtmlAutomationRequest[] Actions { get; set; } = Array.Empty<HtmlAutomationRequest>();
    public string? FinalReadyExpression { get; set; }
    public long MaxOutputBytesPerArtifact { get; set; } = 8L * 1024 * 1024;
    public long MaxTotalOutputBytes { get; set; } = 12L * 1024 * 1024;
}

internal sealed class HtmlPublicResourceBatch {
    public HtmlRuntimeResource[] Resources { get; set; } = Array.Empty<HtmlRuntimeResource>();
    public HtmlRuntimeFetchReplay[] FetchReplays { get; set; } = Array.Empty<HtmlRuntimeFetchReplay>();
    public HtmlRuntimeNavigationReplay[] NavigationReplays { get; set; } = Array.Empty<HtmlRuntimeNavigationReplay>();
}

internal static class HtmlRuntimeNavigationTranscript {
    internal static void Validate(IReadOnlyList<HtmlRuntimeNavigationReplay> supplied,
        IReadOnlyList<string> consumedIdentities) {
        ArgumentNullException.ThrowIfNull(supplied);
        ArgumentNullException.ThrowIfNull(consumedIdentities);
        string[] expected = supplied.Select(replay => replay?.Identity
            ?? throw new HtmlScriptRuntimeException("The navigation replay transcript contains a null replay.")).ToArray();
        if (!expected.SequenceEqual(consumedIdentities, StringComparer.Ordinal))
            throw new HtmlScriptRuntimeException("The isolated execution did not consume the acquired navigation replay transcript exactly.");
    }
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
    public HtmlPublicFontPackageIdentity? FontPackage { get; set; }
    public bool DiscoveryComplete { get; set; }
    public string[] DiscoveryUrls { get; set; } = Array.Empty<string>();
    public HtmlRuntimeFetchDiscovery[] DiscoveryRequests { get; set; } = Array.Empty<HtmlRuntimeFetchDiscovery>();
    public HtmlRuntimeNavigationDiscovery[] NavigationRequests { get; set; } = Array.Empty<HtmlRuntimeNavigationDiscovery>();
    public string[] ConsumedNavigationReplayIdentities { get; set; } = Array.Empty<string>();
    public string? ErrorKind { get; set; }
    public string? Error { get; set; }
    public string? ProviderId { get; set; }
    public string? CaptureUrl { get; set; }
    public string? CaptureManifest { get; set; }
    public string[] TraceEntries { get; set; } = Array.Empty<string>();
    public HtmlAutomationResult[] Actions { get; set; } = Array.Empty<HtmlAutomationResult>();
    public byte[]? Screen { get; set; }
    public byte[]? Print { get; set; }
    public byte[]? ScreenToPage { get; set; }
}
