namespace OfficeIMO.Markdown;

/// <summary>
/// Diagnostic emitted when a roundtrip writer cannot preserve source bytes losslessly.
/// </summary>
public sealed class MarkdownRoundtripDiagnostic {
    /// <summary>Creates a roundtrip diagnostic.</summary>
    public MarkdownRoundtripDiagnostic(string id, string message) {
        Id = id ?? string.Empty;
        Message = message ?? string.Empty;
    }

    /// <summary>Stable diagnostic identifier.</summary>
    public string Id { get; }

    /// <summary>Human-readable diagnostic message.</summary>
    public string Message { get; }
}
