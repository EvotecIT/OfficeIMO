namespace OfficeIMO.Markdown;

/// <summary>
/// Controls how raw HTML blocks (and HTML comments) are emitted during HTML rendering.
/// </summary>
public enum RawHtmlHandling {
    /// <summary>Emit raw HTML verbatim (unsafe for untrusted input).</summary>
    Allow = 0,
    /// <summary>Escape raw HTML so it is displayed as text.</summary>
    Escape = 1,
    /// <summary>Drop raw HTML blocks entirely.</summary>
    Strip = 2
}

