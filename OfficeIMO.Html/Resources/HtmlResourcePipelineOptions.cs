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

    /// <summary>Additional responsive image candidate cap per source set. Null adds no cap; a stricter shared limit still applies.</summary>
    public int? MaxResponsiveImageCandidates { get; set; } = HtmlConversionLimits.DefaultMaxResponsiveImageCandidates;

    /// <summary>Additional character cap for one responsive image <c>sizes</c> value. A stricter shared limit still applies.</summary>
    public int? MaxResponsiveImageSizesCharacters { get; set; } = HtmlConversionLimits.DefaultMaxResponsiveImageSizesCharacters;

    /// <summary>CSS media context used when deciding whether media-gated resources are active.</summary>
    public HtmlCssMediaContext MediaContext { get; set; } = HtmlCssMediaContext.Screen;

    /// <summary>Optional media-query surface width in CSS pixels. When omitted, the context default is used.</summary>
    public double? MediaWidth { get; set; }

    /// <summary>Optional media-query surface height in CSS pixels. When omitted, the context default is used.</summary>
    public double? MediaHeight { get; set; }

    /// <summary>Target device-pixel density used to select one responsive image candidate.</summary>
    public double DevicePixelRatio { get; set; } = 1D;

    /// <summary>Root and initial font size used to resolve relative lengths in responsive image <c>sizes</c>.</summary>
    public double DefaultFontSize { get; set; } = 16D;

    /// <summary>Static device and user-preference values used when deciding whether media-gated resources are active.</summary>
    public HtmlRenderMediaFeatures MediaFeatures { get; set; } = new HtmlRenderMediaFeatures();

    // Generic discovery keeps document icons; rendering omits these non-visual metadata assets.
    internal bool IncludeDocumentIcons { get; set; } = true;
}
