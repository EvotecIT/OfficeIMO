using OfficeIMO.Html.Dom;

namespace OfficeIMO.Html.Runtime;

/// <summary>A local scripted document supplied by a trusted caller. Network loading is not part of this profile.</summary>
public sealed class HtmlScriptRequest {
    /// <summary>Complete HTML source. Inline classic scripts execute while the document loads.</summary>
    public string Html { get; set; } = string.Empty;
    /// <summary>Classic scripts executed in order after document loading.</summary>
    public IReadOnlyList<string> Scripts { get; set; } = Array.Empty<string>();
    /// <summary>A JavaScript expression which must evaluate to boolean true before capture.</summary>
    public string ReadyExpression { get; set; } = "true";
    /// <summary>Deadline per session command, including queue admission and capture transfer; total deadline for CaptureTrustedAsync.</summary>
    public TimeSpan Timeout { get; set; } = TimeSpan.FromSeconds(10);
    /// <summary>Maximum live worker lifetime, including idle time, from startup until termination.</summary>
    public TimeSpan SessionTimeout { get; set; } = TimeSpan.FromMinutes(5);
    /// <summary>Interval between readiness checks. This does not imply network or layout stability.</summary>
    public TimeSpan PollInterval { get; set; } = TimeSpan.FromMilliseconds(10);
    /// <summary>Combined UTF-16 source budget for HTML, supplied scripts and the readiness expression.</summary>
    public int MaxInputCharacters { get; set; } = 8 * 1024 * 1024;
    /// <summary>UTF-16 budget for the complete serialized worker response.</summary>
    public int MaxOutputCharacters { get; set; } = 8 * 1024 * 1024;
    /// <summary>Maximum attached nodes in the captured document, including template fragments.</summary>
    public int MaxNodes { get; set; } = 100_000;
    /// <summary>Maximum captured element nesting depth.</summary>
    public int MaxDepth { get; set; } = 256;
    /// <summary>Maximum tracked promise rejections awaiting a handler within one script turn.</summary>
    public int MaxPendingPromiseRejections { get; set; } = 1024;

    internal HtmlScriptRequest Snapshot() {
        if (Html == null || Scripts == null || ReadyExpression == null) throw new ArgumentException("HTML, scripts and readiness are required.");
        if (Timeout <= TimeSpan.Zero || Timeout > TimeSpan.FromMinutes(5)) throw new ArgumentOutOfRangeException(nameof(Timeout));
        if (SessionTimeout <= TimeSpan.Zero || SessionTimeout > TimeSpan.FromHours(1)) throw new ArgumentOutOfRangeException(nameof(SessionTimeout));
        if (PollInterval < TimeSpan.FromMilliseconds(1) || PollInterval > Timeout) throw new ArgumentOutOfRangeException(nameof(PollInterval));
        if (MaxInputCharacters <= 0 || MaxOutputCharacters <= 0 || MaxNodes <= 0 || MaxDepth <= 0 || MaxPendingPromiseRejections <= 0) throw new ArgumentOutOfRangeException(nameof(MaxInputCharacters), "Resource limits must be positive.");
        var scripts = Scripts.ToArray();
        long length = (long)Html.Length + ReadyExpression.Length;
        foreach (string script in scripts) {
            if (script == null) throw new ArgumentException("A script cannot be null.", nameof(Scripts));
            length += script.Length;
        }
        if (length > MaxInputCharacters) throw new ArgumentException("The combined script input exceeds MaxInputCharacters.");
        return new HtmlScriptRequest { Html = Html, Scripts = scripts, ReadyExpression = ReadyExpression, Timeout = Timeout,
            SessionTimeout = SessionTimeout, PollInterval = PollInterval, MaxInputCharacters = MaxInputCharacters, MaxOutputCharacters = MaxOutputCharacters,
            MaxNodes = MaxNodes, MaxDepth = MaxDepth, MaxPendingPromiseRejections = MaxPendingPromiseRejections };
    }
}

/// <summary>Optional execution provider. Implementations return independent snapshots without exposing interpreter objects.</summary>
public interface IHtmlScriptRuntimeProvider {
    /// <summary>Opens a persistent trusted local document after loading HTML and executing supplied scripts.</summary>
    Task<IHtmlRuntimeSession> OpenTrustedAsync(HtmlScriptRequest request, CancellationToken cancellationToken = default);
    /// <summary>Executes trusted local content and captures it when the explicit readiness condition holds.</summary>
    Task<HtmlScriptCapture> CaptureTrustedAsync(HtmlScriptRequest request, CancellationToken cancellationToken = default);
}

/// <summary>A completed scripted-document capture. Subsequent inspection and conversion are inert.</summary>
public sealed class HtmlScriptCapture {
    /// <summary>Creates a capture from an independent frozen document supplied by a runtime provider.</summary>
    public HtmlScriptCapture(HtmlDocument document, string providerId) {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentException.ThrowIfNullOrWhiteSpace(providerId);
        if (!document.IsReadOnly) throw new ArgumentException("A captured document must be frozen.", nameof(document));
        Document = document;
        ProviderId = providerId;
    }
    /// <summary>Frozen owned document, including structural DOM mutations and template contents.</summary>
    public HtmlDocument Document { get; }
    /// <summary>Actual runtime implementation and version used by the worker.</summary>
    public string ProviderId { get; }
}

/// <summary>Script execution, worker protocol or capture failed. No partial document is returned.</summary>
public sealed class HtmlScriptRuntimeException : InvalidOperationException {
    /// <summary>Creates a runtime failure with a provider-independent message.</summary>
    public HtmlScriptRuntimeException(string message) : base(message) { }
}
