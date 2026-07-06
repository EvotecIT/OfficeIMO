namespace OfficeIMO.Pdf;

/// <summary>
/// Describes one failed redaction verification check.
/// </summary>
public sealed class PdfRedactionVerificationIssue {
    internal PdfRedactionVerificationIssue(string feature, string marker, string message) {
        Feature = feature;
        Marker = marker;
        Message = message;
    }

    /// <summary>Stable feature name, for example RemovedTextMarker or RetainedTextMarker.</summary>
    public string Feature { get; }

    /// <summary>Marker text involved in the verification issue.</summary>
    public string Marker { get; }

    /// <summary>Human-readable redaction verification diagnostic.</summary>
    public string Message { get; }
}
