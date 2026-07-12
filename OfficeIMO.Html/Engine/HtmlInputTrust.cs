namespace OfficeIMO.Html;

/// <summary>
/// Declares the caller-assigned trust boundary for HTML input independently from conversion fidelity.
/// </summary>
public enum HtmlInputTrust {
    /// <summary>
    /// Input is not trusted; adapters should use bounded, offline-safe resource defaults.
    /// </summary>
    Untrusted,

    /// <summary>
    /// Input is trusted by the caller; adapters may enable document-provided resources within their configured policies.
    /// </summary>
    Trusted
}
