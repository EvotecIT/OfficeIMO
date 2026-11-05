using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Markdown;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Creates native Markdown HTML document, styling and resource-link settings.</summary>
[Cmdlet(VerbsCommon.New, "OfficeMarkdownHtmlOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(MarkdownHtmlExportOptions))]
public sealed class NewOfficeMarkdownHtmlOptionsCommand : OfficeCmdlet {
    /// <summary>Native syntax, normalization, URL and input-limit policy, captured independently.</summary>
    [Parameter(ParameterSetName = "Default"), OfficeOptionsTransformation(typeof(MarkdownLoadOptions))]
    public MarkdownLoadOptions? ReaderOptions { get; set; }
    /// <summary>Built-in theme, shared theme object or theme template.</summary>
    [Parameter(ParameterSetName = "Default"), OfficeThemeTransformation]
    public OfficeThemeInput? Theme { get; set; }
    /// <summary>Render a complete HTML document; explicit false renders a fragment.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter DocumentMode { get; set; }
    /// <summary>Native Style setting.</summary>
    [Parameter]
    public MarkdownHtmlStyle? Style { get; set; }
    /// <summary>Native CssDelivery setting.</summary>
    [Parameter]
    public MarkdownCssDelivery? CssDelivery { get; set; }
    /// <summary>Native AssetMode setting.</summary>
    [Parameter]
    public MarkdownAssetMode? AssetMode { get; set; }
    /// <summary>Native Title setting.</summary>
    [Parameter]
    public string? Title { get; set; }
    /// <summary>Native RawHtmlHandling setting.</summary>
    [Parameter]
    public MarkdownRawHtmlHandling? RawHtmlHandling { get; set; }
    /// <summary>Native ExternalLinksRel setting.</summary>
    [Parameter]
    public string? ExternalLinksRel { get; set; }
    /// <summary>Native ExternalLinksReferrerPolicy setting.</summary>
    [Parameter]
    public string? ExternalLinksReferrerPolicy { get; set; }
    /// <summary>Native ImagesReferrerPolicy setting.</summary>
    [Parameter]
    public string? ImagesReferrerPolicy { get; set; }
    /// <summary>Enable or explicitly disable IncludeAnchorLinks.</summary>
    [Parameter]
    public SwitchParameter IncludeAnchorLinks { get; set; }
    /// <summary>Enable or explicitly disable GitHubTaskListHtml.</summary>
    [Parameter]
    public SwitchParameter GitHubTaskListHtml { get; set; }
    /// <summary>Enable or explicitly disable GitHubFootnoteHtml.</summary>
    [Parameter]
    public SwitchParameter GitHubFootnoteHtml { get; set; }
    /// <summary>Enable or explicitly disable ExternalLinksTargetBlank.</summary>
    [Parameter]
    public SwitchParameter ExternalLinksTargetBlank { get; set; }
    /// <summary>Enable or explicitly disable RestrictHttpLinksToBaseOrigin.</summary>
    [Parameter]
    public SwitchParameter RestrictHttpLinksToBaseOrigin { get; set; }
    /// <summary>Enable or explicitly disable RestrictHttpImagesToBaseOrigin.</summary>
    [Parameter]
    public SwitchParameter RestrictHttpImagesToBaseOrigin { get; set; }
    /// <summary>Enable or explicitly disable BlockExternalHttpImages.</summary>
    [Parameter]
    public SwitchParameter BlockExternalHttpImages { get; set; }
    /// <summary>Enable or explicitly disable ImagesLoadingLazy.</summary>
    [Parameter]
    public SwitchParameter ImagesLoadingLazy { get; set; }
    /// <summary>Enable or explicitly disable ImagesDecodingAsync.</summary>
    [Parameter]
    public SwitchParameter ImagesDecodingAsync { get; set; }
    /// <summary>Additional allowed HTTP hosts.</summary>
    [Parameter]
    public string[]? AllowedHttpLinkHost { get; set; }
    /// <summary>Additional allowed HTTP hosts.</summary>
    [Parameter]
    public string[]? AllowedHttpImageHost { get; set; }
    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new MarkdownHtmlExportOptions { ReaderOptions = ReaderOptions?.Clone() };
        if (Theme != null) options.Theme = MarkdownVisualTheme.FromOfficeTheme(Theme.Resolve(OfficeTheme.FromDocument));
        if (BoundSwitch(nameof(DocumentMode)) is bool document) options.Kind = document ? MarkdownHtmlKind.Document : MarkdownHtmlKind.Fragment;
        if (Style.HasValue) options.Style = Style.Value;
        if (CssDelivery.HasValue) options.CssDelivery = CssDelivery.Value;
        if (AssetMode.HasValue) options.AssetMode = AssetMode.Value;
        if (Title != null) options.Title = Title;
        if (RawHtmlHandling.HasValue) options.RawHtmlHandling = RawHtmlHandling.Value;
        if (ExternalLinksRel != null) options.ExternalLinksRel = ExternalLinksRel;
        if (ExternalLinksReferrerPolicy != null) options.ExternalLinksReferrerPolicy = ExternalLinksReferrerPolicy;
        if (ImagesReferrerPolicy != null) options.ImagesReferrerPolicy = ImagesReferrerPolicy;
        if (BoundSwitch(nameof(IncludeAnchorLinks)) is bool includeAnchorLinks) options.IncludeAnchorLinks = includeAnchorLinks;
        if (BoundSwitch(nameof(GitHubTaskListHtml)) is bool gitHubTaskListHtml) options.GitHubTaskListHtml = gitHubTaskListHtml;
        if (BoundSwitch(nameof(GitHubFootnoteHtml)) is bool gitHubFootnoteHtml) options.GitHubFootnoteHtml = gitHubFootnoteHtml;
        if (BoundSwitch(nameof(ExternalLinksTargetBlank)) is bool externalLinksTargetBlank) options.ExternalLinksTargetBlank = externalLinksTargetBlank;
        if (BoundSwitch(nameof(RestrictHttpLinksToBaseOrigin)) is bool restrictHttpLinksToBaseOrigin) options.RestrictHttpLinksToBaseOrigin = restrictHttpLinksToBaseOrigin;
        if (BoundSwitch(nameof(RestrictHttpImagesToBaseOrigin)) is bool restrictHttpImagesToBaseOrigin) options.RestrictHttpImagesToBaseOrigin = restrictHttpImagesToBaseOrigin;
        if (BoundSwitch(nameof(BlockExternalHttpImages)) is bool blockExternalHttpImages) options.BlockExternalHttpImages = blockExternalHttpImages;
        if (BoundSwitch(nameof(ImagesLoadingLazy)) is bool imagesLoadingLazy) options.ImagesLoadingLazy = imagesLoadingLazy;
        if (BoundSwitch(nameof(ImagesDecodingAsync)) is bool imagesDecodingAsync) options.ImagesDecodingAsync = imagesDecodingAsync;
        if (AllowedHttpLinkHost != null) options.AllowedHttpLinkHosts.AddRange(AllowedHttpLinkHost);
        if (AllowedHttpImageHost != null) options.AllowedHttpImageHosts.AddRange(AllowedHttpImageHost);
        WriteObject(options);
    }
}
