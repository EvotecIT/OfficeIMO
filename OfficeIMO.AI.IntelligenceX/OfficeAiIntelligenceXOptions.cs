namespace OfficeIMO.AI.IntelligenceX;

/// <summary>Restricted SDK connection modes for document evidence.</summary>
public enum OfficeAiIntelligenceXTransport {
    /// <summary>IX's native ChatGPT transport using its configured local authentication store.</summary>
    ChatGpt,
    /// <summary>Explicit OpenAI-compatible HTTP endpoint, hosted or loopback-local.</summary>
    CompatibleHttp,
    /// <summary>Native Copilot HTTP requests through the shared IntelligenceX client.</summary>
    CopilotNative
}

/// <summary>Connection-only settings; document operations do not depend on these transport details.</summary>
public sealed class OfficeAiIntelligenceXOptions {
    /// <summary>Selected restricted transport.</summary>
    public OfficeAiIntelligenceXTransport Transport { get; init; } = OfficeAiIntelligenceXTransport.ChatGpt;
    /// <summary>Explicit API base URL for CompatibleHttp, including its version prefix when needed.</summary>
    public Uri? Endpoint { get; init; }
    /// <summary>Caller-supplied API credential; never included in results or diagnostics.</summary>
    public string? ApiKey { get; init; }
    /// <summary>Optional host-owned Copilot credential and connection settings. The SDK owns authentication and HTTP behavior.</summary>
    public global::IntelligenceX.Copilot.Native.CopilotNativeOptions? CopilotOptions { get; init; }
    /// <summary>Whether the compatible endpoint supports SSE. Non-streaming retains the same result contract.</summary>
    public bool Streaming { get; init; }
    /// <summary>Explicitly prefer the current local Codex login over IX's saved ChatGPT credential.</summary>
    public bool PreferCurrentCodexSession { get; init; }
}
