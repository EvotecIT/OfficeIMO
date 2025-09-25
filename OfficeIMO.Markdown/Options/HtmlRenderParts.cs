namespace OfficeIMO.Markdown;

/// <summary>Parts of HTML output for advanced embedding.</summary>
public sealed class HtmlRenderParts {
    /// <summary>HTML inside &lt;head&gt; (excluding outer tag).</summary>
    public string Head { get; set; } = string.Empty;
    /// <summary>Rendered content inside &lt;body&gt; (excluding outer tag). For fragments, this is the whole HTML.</summary>
    public string Body { get; set; } = string.Empty;
    /// <summary>CSS content (inline text) that was used/generated. Useful when <see cref="CssDelivery.ExternalFile"/>.</summary>
    public string Css { get; set; } = string.Empty;
    /// <summary>Inline scripts emitted by the renderer, if any.</summary>
    public string Scripts { get; set; } = string.Empty;
    /// <summary>Declared assets (CSS/JS) with stable ids so hosts can deduplicate across fragments/documents.</summary>
    public System.Collections.Generic.List<HtmlAsset> Assets { get; set; } = new();
}

