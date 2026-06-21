namespace OfficeIMO.Html;

/// <summary>
/// Options controlling shared OfficeIMO HTML resource planning.
/// </summary>
public sealed class HtmlResourcePipelineOptions {
    /// <summary>Optional base URI used to resolve relative resource references.</summary>
    public Uri? BaseUri { get; set; }

    /// <summary>URL policy applied before resource references are reported as allowed.</summary>
    public HtmlUrlPolicy UrlPolicy { get; set; } = HtmlUrlPolicy.CreateOfficeIMOProfile();

    /// <summary>Maximum responsive image candidates considered per source set. Null means unbounded.</summary>
    public int? MaxResponsiveImageCandidates { get; set; } = 32;

    /// <summary>CSS media context used when deciding whether media-gated resources are active.</summary>
    public HtmlCssMediaContext MediaContext { get; set; } = HtmlCssMediaContext.Screen;
}
