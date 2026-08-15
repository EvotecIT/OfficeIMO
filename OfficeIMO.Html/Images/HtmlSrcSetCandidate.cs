namespace OfficeIMO.Html;

/// <summary>
/// Represents one candidate in an HTML <c>srcset</c> attribute.
/// </summary>
public readonly struct HtmlSrcSetCandidate {
    /// <summary>
    /// Creates a source-set candidate.
    /// </summary>
    public HtmlSrcSetCandidate(string url, string descriptor) {
        UrlStart = -1;
        Url = url ?? string.Empty;
        Descriptor = descriptor ?? string.Empty;
    }

    internal HtmlSrcSetCandidate(string url, string descriptor, int urlStart) {
        Url = url ?? string.Empty;
        Descriptor = descriptor ?? string.Empty;
        UrlStart = urlStart;
    }

    /// <summary>
    /// Candidate URL before caller-specific resolution.
    /// </summary>
    public string Url { get; }

    /// <summary>
    /// Candidate descriptor, such as <c>2x</c> or <c>640w</c>.
    /// </summary>
    public string Descriptor { get; }

    /// <summary>
    /// Zero-based offset of <see cref="Url"/> in the parsed source-set value, or <c>-1</c> when the candidate was constructed directly.
    /// </summary>
    public int UrlStart { get; }
}
