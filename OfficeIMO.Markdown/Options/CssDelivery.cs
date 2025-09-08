namespace OfficeIMO.Markdown;

/// <summary>
/// How CSS is provided to the rendered HTML.
/// </summary>
public enum CssDelivery {
    /// <summary>Inline a &lt;style&gt; tag with CSS content.</summary>
    Inline,
    /// <summary>Write CSS to a sidecar file and link it from HTML.</summary>
    ExternalFile,
    /// <summary>Use a &lt;link rel="stylesheet" href="..."&gt; tag to a provided URL.</summary>
    LinkHref,
    /// <summary>Do not include any CSS.</summary>
    None
}

