namespace OfficeIMO.Html;

/// <summary>
/// Options for building the shared OfficeIMO HTML conversion document.
/// </summary>
public sealed class HtmlConversionDocumentOptions {
    /// <summary>Conversion profile the document should advertise to downstream adapters.</summary>
    public HtmlConversionProfile Profile { get; set; } = HtmlConversionProfile.Semantic;

    /// <summary>
    /// Caller-assigned input trust boundary. This is independent from the semantic or visual fidelity profile.
    /// </summary>
    public HtmlInputTrust Trust { get; set; } = HtmlInputTrust.Untrusted;

    /// <summary>Optional base URI used for URL normalization and resource planning.</summary>
    public Uri? BaseUri { get; set; }

    /// <summary>URL policy used for resource planning and normalized HTML output.</summary>
    public HtmlUrlPolicy UrlPolicy { get; set; } = HtmlUrlPolicy.CreateOfficeIMOProfile();

    /// <summary>When true, logical analysis and normalized output use body contents as the conversion root.</summary>
    public bool UseBodyContentsOnly { get; set; } = true;

    /// <summary>When true, the builder emits a normalized HTML representation alongside the source HTML.</summary>
    public bool IncludeNormalizedHtml { get; set; } = true;

    /// <summary>Options used when producing normalized HTML.</summary>
    public HtmlNormalizationOptions NormalizationOptions { get; set; } = new HtmlNormalizationOptions();

    /// <summary>Creates a resource-pipeline options snapshot from this conversion-document configuration.</summary>
    public HtmlResourcePipelineOptions ToResourcePipelineOptions() {
        return new HtmlResourcePipelineOptions {
            BaseUri = BaseUri,
            UrlPolicy = (UrlPolicy ?? HtmlUrlPolicy.CreateOfficeIMOProfile()).Clone(),
            MaxResponsiveImageCandidates = null,
            MediaContext = Profile == HtmlConversionProfile.HighFidelityPrint
                ? HtmlCssMediaContext.Print
                : HtmlCssMediaContext.Screen
        };
    }

    /// <summary>Creates an independent options snapshot that can be safely adjusted for one load operation.</summary>
    public HtmlConversionDocumentOptions Clone() {
        HtmlNormalizationOptions normalization = NormalizationOptions ?? new HtmlNormalizationOptions();
        return new HtmlConversionDocumentOptions {
            Profile = Profile,
            Trust = Trust,
            BaseUri = BaseUri,
            UrlPolicy = (UrlPolicy ?? HtmlUrlPolicy.CreateOfficeIMOProfile()).Clone(),
            UseBodyContentsOnly = UseBodyContentsOnly,
            IncludeNormalizedHtml = IncludeNormalizedHtml,
            NormalizationOptions = new HtmlNormalizationOptions {
                BaseUri = normalization.BaseUri,
                BaseElementBaseUri = normalization.BaseElementBaseUri,
                UrlPolicy = (normalization.UrlPolicy ?? HtmlUrlPolicy.CreateOfficeIMOProfile()).Clone(),
                UseBodyContentsOnly = normalization.UseBodyContentsOnly,
                PreserveComments = normalization.PreserveComments,
                PreserveSkippedElementMarkers = normalization.PreserveSkippedElementMarkers,
                PreserveStyleElements = normalization.PreserveStyleElements,
                RemoveEventHandlerAttributes = normalization.RemoveEventHandlerAttributes,
                CollapseTextWhitespace = normalization.CollapseTextWhitespace
            }
        };
    }
}
