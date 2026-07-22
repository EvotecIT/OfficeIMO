namespace OfficeIMO.Html;

/// <summary>
/// Options for building the shared OfficeIMO HTML conversion document.
/// </summary>
public sealed class HtmlConversionDocumentOptions {
    private HtmlInputTrust _trust = HtmlInputTrust.Untrusted;
    private HtmlUrlPolicy _urlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile();
    private HtmlUrlPolicy _resourceUrlPolicy = CreateResourcePolicy(HtmlUrlPolicy.CreateWebOnlyProfile());
    private HtmlConversionLimits _limits = HtmlConversionLimits.CreateUntrustedProfile();
    private bool _urlPolicyExplicit;
    private bool _resourceUrlPolicyExplicit;
    private bool _limitsExplicit;

    /// <summary>Creates bounded, web-only defaults for untrusted HTML.</summary>
    public static HtmlConversionDocumentOptions CreateUntrustedProfile() => new HtmlConversionDocumentOptions();

    /// <summary>Creates compatibility-oriented defaults for caller-trusted HTML.</summary>
    public static HtmlConversionDocumentOptions CreateTrustedProfile() => new HtmlConversionDocumentOptions {
        Trust = HtmlInputTrust.Trusted,
        UrlPolicy = HtmlUrlPolicy.CreateOfficeIMOProfile(),
        Limits = HtmlConversionLimits.CreateTrustedProfile()
    };

    /// <summary>Conversion profile the document should advertise to downstream adapters.</summary>
    public HtmlConversionProfile Profile { get; set; } = HtmlConversionProfile.Semantic;

    /// <summary>
    /// Caller-assigned input trust boundary. This is independent from the semantic or visual fidelity profile.
    /// </summary>
    public HtmlInputTrust Trust {
        get => _trust;
        set {
            _trust = value;
            if (!_urlPolicyExplicit) _urlPolicy = ResolveDefaultUrlPolicy(value);
            if (!_resourceUrlPolicyExplicit) _resourceUrlPolicy = CreateResourcePolicy(_urlPolicy);
            if (!_limitsExplicit) _limits = ResolveDefaultLimits(value);
        }
    }

    /// <summary>Optional base URI used for URL normalization and resource planning.</summary>
    public Uri? BaseUri { get; set; }

    /// <summary>URL policy used for hyperlinks, navigation targets, and normalized HTML output.</summary>
    public HtmlUrlPolicy UrlPolicy {
        get => _urlPolicy;
        set {
            _urlPolicy = value ?? ResolveDefaultUrlPolicy(Trust);
            _urlPolicyExplicit = value != null;
            if (!_resourceUrlPolicyExplicit) _resourceUrlPolicy = CreateResourcePolicy(_urlPolicy);
        }
    }

    /// <summary>Separate URL policy for images, stylesheets, fonts, media, and other non-navigation resources.</summary>
    public HtmlUrlPolicy ResourceUrlPolicy {
        get => _resourceUrlPolicy;
        set {
            _resourceUrlPolicy = value ?? ResolveDefaultResourceUrlPolicy(Trust);
            _resourceUrlPolicyExplicit = value != null;
        }
    }

    /// <summary>Shared limits applied before logical, style, resource, normalization, or adapter work.</summary>
    public HtmlConversionLimits Limits {
        get => _limits;
        set {
            _limits = value ?? ResolveDefaultLimits(Trust);
            _limitsExplicit = value != null;
        }
    }

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
            UrlPolicy = (UrlPolicy ?? ResolveDefaultUrlPolicy(Trust)).Clone(),
            ResourceUrlPolicy = (ResourceUrlPolicy ?? ResolveDefaultResourceUrlPolicy(Trust)).Clone(),
            Limits = (Limits ?? HtmlConversionLimits.CreateUntrustedProfile()).Clone(),
            MaxResponsiveImageCandidates = (Limits ?? HtmlConversionLimits.CreateUntrustedProfile()).MaxResponsiveImageCandidates,
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
            ResourceUrlPolicy = (ResourceUrlPolicy ?? ResolveDefaultResourceUrlPolicy(Trust)).Clone(),
            Limits = (Limits ?? ResolveDefaultLimits(Trust)).Clone(),
            UseBodyContentsOnly = UseBodyContentsOnly,
            IncludeNormalizedHtml = IncludeNormalizedHtml,
            NormalizationOptions = new HtmlNormalizationOptions {
                BaseUri = normalization.BaseUri,
                BaseElementBaseUri = normalization.BaseElementBaseUri,
                UrlPolicy = (normalization.UrlPolicy ?? HtmlUrlPolicy.CreateOfficeIMOProfile()).Clone(),
                ResourceUrlPolicy = normalization.ResourceUrlPolicy?.Clone(),
                Limits = (normalization.Limits ?? ResolveDefaultLimits(Trust)).Clone(),
                MaxResponsiveImageCandidates = normalization.MaxResponsiveImageCandidates,
                UseBodyContentsOnly = normalization.UseBodyContentsOnly,
                PreserveComments = normalization.PreserveComments,
                PreserveSkippedElementMarkers = normalization.PreserveSkippedElementMarkers,
                PreserveStyleElements = normalization.PreserveStyleElements,
                RemoveEventHandlerAttributes = normalization.RemoveEventHandlerAttributes,
                CollapseTextWhitespace = normalization.CollapseTextWhitespace
            }
        };
    }

    internal void Validate() {
        if (!Enum.IsDefined(typeof(HtmlInputTrust), Trust)) throw new ArgumentOutOfRangeException(nameof(Trust));
        if (!Enum.IsDefined(typeof(HtmlConversionProfile), Profile)) throw new ArgumentOutOfRangeException(nameof(Profile));
        (Limits ?? ResolveDefaultLimits(Trust)).Validate();
    }

    private static HtmlConversionLimits ResolveDefaultLimits(HtmlInputTrust trust) =>
        trust == HtmlInputTrust.Trusted
            ? HtmlConversionLimits.CreateTrustedProfile()
            : HtmlConversionLimits.CreateUntrustedProfile();

    private static HtmlUrlPolicy ResolveDefaultUrlPolicy(HtmlInputTrust trust) =>
        trust == HtmlInputTrust.Trusted
            ? HtmlUrlPolicy.CreateOfficeIMOProfile()
            : HtmlUrlPolicy.CreateWebOnlyProfile();

    private static HtmlUrlPolicy ResolveDefaultResourceUrlPolicy(HtmlInputTrust trust) =>
        CreateResourcePolicy(ResolveDefaultUrlPolicy(trust));

    private static HtmlUrlPolicy CreateResourcePolicy(HtmlUrlPolicy hyperlinkPolicy) {
        HtmlUrlPolicy resourcePolicy = HtmlResourceUrlPolicy.Create(hyperlinkPolicy);
        // Data URLs are not safe navigation targets, but bounded embedded resources are a core
        // offline conversion path. External HTTP(S) resources still require an explicit resolver.
        resourcePolicy.AllowDataUrls = true;
        if (resourcePolicy.RestrictUrlSchemes) resourcePolicy.AllowedUrlSchemes.Add("data");
        return resourcePolicy;
    }
}
