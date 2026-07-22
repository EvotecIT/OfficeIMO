namespace OfficeIMO.Html;

/// <summary>
/// Options controlling shared OfficeIMO HTML resource planning.
/// </summary>
public sealed class HtmlResourcePipelineOptions {
    /// <summary>Optional base URI used to resolve relative resource references.</summary>
    public Uri? BaseUri { get; set; }

    /// <summary>URL policy applied before resource references are reported as allowed.</summary>
    public HtmlUrlPolicy UrlPolicy { get; set; } = HtmlUrlPolicy.CreateOfficeIMOProfile();

    /// <summary>
    /// Optional separate policy for images, stylesheets, fonts, media, and other non-hyperlink resources.
    /// When omitted, it is derived from <see cref="UrlPolicy"/>.
    /// </summary>
    public HtmlUrlPolicy? ResourceUrlPolicy { get; set; }

    /// <summary>Shared source, DOM, stylesheet, and semantic-metadata limits applied before discovery.</summary>
    public HtmlConversionLimits Limits { get; set; } = HtmlConversionLimits.CreateUntrustedProfile();

    /// <summary>Maximum responsive image candidates considered per source set. Null means unbounded.</summary>
    public int? MaxResponsiveImageCandidates { get; set; } = HtmlConversionLimits.DefaultMaxResponsiveImageCandidates;

    /// <summary>CSS media context used when deciding whether media-gated resources are active.</summary>
    public HtmlCssMediaContext MediaContext { get; set; } = HtmlCssMediaContext.Screen;

    /// <summary>Optional media-query surface width in CSS pixels. When omitted, the context default is used.</summary>
    public double? MediaWidth { get; set; }

    /// <summary>Optional media-query surface height in CSS pixels. When omitted, the context default is used.</summary>
    public double? MediaHeight { get; set; }
}
