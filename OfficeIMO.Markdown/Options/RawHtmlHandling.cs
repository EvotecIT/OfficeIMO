namespace OfficeIMO.Markdown;

/// <summary>
/// Controls how raw HTML blocks (and HTML comments) are emitted during HTML rendering.
/// </summary>
public enum RawHtmlHandling {
    /// <summary>Emit raw HTML verbatim (unsafe for untrusted input).</summary>
    Allow = 0,
    /// <summary>Escape raw HTML so it is displayed as text.</summary>
    Escape = 1,
    /// <summary>
    /// Sanitize raw HTML using a small allowlist. Disallowed tags are escaped and dangerous attributes are removed.
    /// Intended for pragmatic "HTML on but safe-ish" scenarios, not full HTML policy enforcement.
    /// </summary>
    Sanitize = 2,
    /// <summary>Drop raw HTML blocks entirely.</summary>
    Strip = 3
}
