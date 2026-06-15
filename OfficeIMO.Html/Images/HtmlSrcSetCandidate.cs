namespace OfficeIMO.Html;

/// <summary>
/// Represents one candidate in an HTML <c>srcset</c> attribute.
/// </summary>
public readonly struct HtmlSrcSetCandidate {
    /// <summary>
    /// Creates a source-set candidate.
    /// </summary>
    public HtmlSrcSetCandidate(string url, string descriptor) {
        Url = url ?? string.Empty;
        Descriptor = descriptor ?? string.Empty;
    }

    /// <summary>
    /// Candidate URL before caller-specific resolution.
    /// </summary>
    public string Url { get; }

    /// <summary>
    /// Candidate descriptor, such as <c>2x</c> or <c>640w</c>.
    /// </summary>
    public string Descriptor { get; }
}
