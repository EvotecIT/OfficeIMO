namespace OfficeIMO.Html;

/// <summary>
/// Immutable rendering intent. It keeps document provenance, CSS media, layout surface,
/// pagination, page selection, and output encoder independent from one another.
/// </summary>
public sealed class HtmlRenderRequest {
    private readonly HtmlRenderOptions _options;

    private HtmlRenderRequest(
        HtmlRenderIntentProfile profile,
        HtmlRenderDocumentState documentState,
        HtmlCssMediaContext cssMedia,
        HtmlRenderLayoutSurface surface,
        HtmlRenderPaginationPolicy pagination,
        HtmlRenderEncoder encoder,
        HtmlRenderPageSet pageSet,
        HtmlRenderOptions options) {
        Profile = profile;
        DocumentState = documentState;
        CssMedia = cssMedia;
        Surface = surface;
        Pagination = pagination;
        Encoder = encoder;
        PageSet = pageSet ?? throw new ArgumentNullException(nameof(pageSet));
        _options = (options ?? throw new ArgumentNullException(nameof(options))).Clone();
        Validate();
    }

    /// <summary>Named versioned profile whose defaults created this request.</summary>
    public HtmlRenderIntentProfile Profile { get; }
    /// <summary>Stable versioned profile identifier.</summary>
    public string ProfileId => HtmlRenderProfileContracts.Get(Profile).Id;
    /// <summary>Whether the effective axes still exactly match the named profile.</summary>
    public bool MatchesNamedProfile {
        get {
            HtmlRenderProfileContract contract = HtmlRenderProfileContracts.Get(Profile);
            return CssMedia == contract.CssMedia && Surface == contract.Surface && Pagination == contract.Pagination;
        }
    }
    /// <summary>Qualification of the effective axis combination. Custom combinations are unqualified.</summary>
    public HtmlCapabilityCoverage Coverage => MatchesNamedProfile
        ? HtmlRenderProfileContracts.Get(Profile).Coverage
        : HtmlCapabilityCoverage.Unqualified;
    /// <summary>Provenance of the immutable document supplied to the operation.</summary>
    public HtmlRenderDocumentState DocumentState { get; }
    /// <summary>CSS media type, independent from pagination and output format.</summary>
    public HtmlCssMediaContext CssMedia { get; }
    /// <summary>Geometry model used to contain the layout.</summary>
    public HtmlRenderLayoutSurface Surface { get; }
    /// <summary>How the layout becomes one or more surfaces.</summary>
    public HtmlRenderPaginationPolicy Pagination { get; }
    /// <summary>Output adapter intended to consume the retained result.</summary>
    public HtmlRenderEncoder Encoder { get; }
    /// <summary>Explicit page selection and composition policy.</summary>
    public HtmlRenderPageSet PageSet { get; }
    /// <summary>Independent rendering, resource, font, safety, and encoder settings snapshot.</summary>
    public HtmlRenderOptions Options => _options.Clone();

    /// <summary>Creates a request from one built-in profile and its declared defaults.</summary>
    public static HtmlRenderRequest Create(
        HtmlRenderIntentProfile profile,
        HtmlRenderEncoder encoder = HtmlRenderEncoder.DisplayList,
        HtmlRenderOptions? options = null,
        HtmlRenderDocumentState documentState = HtmlRenderDocumentState.StaticSource) {
        HtmlRenderProfileContract contract = HtmlRenderProfileContracts.Get(profile);
        HtmlRenderOptions resolved = options?.Clone() ?? CreateDefaultOptions(profile);
        return new HtmlRenderRequest(profile, documentState, contract.CssMedia, contract.Surface,
            contract.Pagination, encoder, contract.DefaultPageSet, resolved);
    }

    /// <summary>Returns a copy selecting another immutable document-state provenance.</summary>
    public HtmlRenderRequest WithDocumentState(HtmlRenderDocumentState state) =>
        new(Profile, state, CssMedia, Surface, Pagination, Encoder, PageSet, _options);

    /// <summary>Returns a copy selecting CSS media independently from layout geometry.</summary>
    public HtmlRenderRequest WithCssMedia(HtmlCssMediaContext cssMedia) =>
        new(Profile, DocumentState, cssMedia, Surface, Pagination, Encoder, PageSet, _options);

    /// <summary>Returns a copy selecting a layout surface independently from CSS media.</summary>
    public HtmlRenderRequest WithLayoutSurface(HtmlRenderLayoutSurface surface) =>
        new(Profile, DocumentState, CssMedia, surface, Pagination, Encoder, PageSet, _options);

    /// <summary>Returns a copy selecting a pagination policy independently from CSS media.</summary>
    public HtmlRenderRequest WithPagination(HtmlRenderPaginationPolicy pagination) =>
        new(Profile, DocumentState, CssMedia, Surface, pagination, Encoder, PageSet, _options);

    /// <summary>Returns a copy changing the coupled surface and pagination geometry atomically.</summary>
    public HtmlRenderRequest WithLayout(
        HtmlRenderLayoutSurface surface,
        HtmlRenderPaginationPolicy pagination) =>
        new(Profile, DocumentState, CssMedia, surface, pagination, Encoder, PageSet, _options);

    /// <summary>Returns a copy changing CSS media, surface, and pagination atomically.</summary>
    public HtmlRenderRequest WithAxes(
        HtmlCssMediaContext cssMedia,
        HtmlRenderLayoutSurface surface,
        HtmlRenderPaginationPolicy pagination) =>
        new(Profile, DocumentState, cssMedia, surface, pagination, Encoder, PageSet, _options);

    /// <summary>Returns a copy selecting another output encoder without changing layout intent.</summary>
    public HtmlRenderRequest WithEncoder(HtmlRenderEncoder encoder) =>
        new(Profile, DocumentState, CssMedia, Surface, Pagination, encoder, PageSet, _options);

    /// <summary>Returns a copy selecting another page-set behavior without rerouting layout intent.</summary>
    public HtmlRenderRequest WithPageSet(HtmlRenderPageSet pageSet) =>
        new(Profile, DocumentState, CssMedia, Surface, Pagination, Encoder, pageSet, _options);

    /// <summary>Returns a copy using a new independent options snapshot.</summary>
    public HtmlRenderRequest WithOptions(HtmlRenderOptions options) =>
        new(Profile, DocumentState, CssMedia, Surface, Pagination, Encoder, PageSet, options);

    internal HtmlRenderOptions ResolveOptions() {
        HtmlRenderOptions resolved = _options.Clone();
        resolved.CssMediaContextOverride = CssMedia;
        resolved.ClipContinuousSurfaceToViewport = Surface == HtmlRenderLayoutSurface.Viewport;
        resolved.Mode = Pagination == HtmlRenderPaginationPolicy.FragmentedReflow
            ? HtmlRenderMode.Paged
            : HtmlRenderMode.Continuous;
        if (Surface == HtmlRenderLayoutSurface.Viewport && !resolved.ViewportHeight.HasValue) {
            throw new ArgumentException("The screen viewport profile requires a finite ViewportHeight.", nameof(Options));
        }
        resolved.Validate();
        return resolved;
    }

    internal static HtmlRenderRequest FromLegacy(
        HtmlRenderOptions? options,
        HtmlRenderEncoder encoder,
        HtmlRenderPageSet pageSet,
        bool forcePrintPaged = false) {
        HtmlRenderOptions resolved = options?.Clone() ?? new HtmlRenderOptions();
        HtmlRenderIntentProfile profile;
        if (forcePrintPaged) {
            profile = HtmlRenderIntentProfile.PrintPaged;
        } else if (resolved.Mode == HtmlRenderMode.Paged && resolved.MediaContext == HtmlCssMediaContext.Screen) {
            profile = HtmlRenderIntentProfile.ScreenMediaPaged;
        } else if (resolved.Mode == HtmlRenderMode.Paged) {
            profile = HtmlRenderIntentProfile.PrintPaged;
        } else {
            profile = HtmlRenderIntentProfile.ScreenFullPage;
        }
        return Create(profile, encoder, resolved).WithPageSet(pageSet);
    }

    private void Validate() {
        if (!Enum.IsDefined(typeof(HtmlRenderDocumentState), DocumentState)) {
            throw new ArgumentOutOfRangeException(nameof(DocumentState));
        }
        if (!Enum.IsDefined(typeof(HtmlCssMediaContext), CssMedia)) {
            throw new ArgumentOutOfRangeException(nameof(CssMedia));
        }
        if (!Enum.IsDefined(typeof(HtmlRenderLayoutSurface), Surface)) {
            throw new ArgumentOutOfRangeException(nameof(Surface));
        }
        if (!Enum.IsDefined(typeof(HtmlRenderPaginationPolicy), Pagination)) {
            throw new ArgumentOutOfRangeException(nameof(Pagination));
        }
        if (!Enum.IsDefined(typeof(HtmlRenderEncoder), Encoder)) {
            throw new ArgumentOutOfRangeException(nameof(Encoder));
        }
        HtmlRenderProfileContract contract = HtmlRenderProfileContracts.Get(Profile);
        if (!contract.Encoders.Contains(Encoder)) {
            throw new NotSupportedException($"Render profile '{contract.Id}' does not admit encoder '{Encoder}'.");
        }
        if (!contract.PageSets.Contains(PageSet.Mode)) {
            throw new NotSupportedException($"Render profile '{contract.Id}' does not admit page-set mode '{PageSet.Mode}'.");
        }
        if (Pagination == HtmlRenderPaginationPolicy.ElementAwarePlacement) {
            throw new NotSupportedException("Element-aware placement is not implemented by the current render profiles.");
        }
        bool unpagedSurface = Surface == HtmlRenderLayoutSurface.Viewport || Surface == HtmlRenderLayoutSurface.Continuous;
        if (unpagedSurface && Pagination != HtmlRenderPaginationPolicy.None) {
            throw new NotSupportedException($"Layout surface '{Surface}' requires pagination 'None'.");
        }
        if (Surface == HtmlRenderLayoutSurface.Paged
            && Pagination != HtmlRenderPaginationPolicy.FragmentedReflow
            && Pagination != HtmlRenderPaginationPolicy.FixedCanvasSlicing) {
            throw new NotSupportedException("A paged layout surface requires fragmented reflow or fixed-canvas slicing.");
        }
    }

    private static HtmlRenderOptions CreateDefaultOptions(HtmlRenderIntentProfile profile) {
        var options = new HtmlRenderOptions();
        if (profile == HtmlRenderIntentProfile.ScreenViewport) {
            options.ViewportHeight = 1056D;
        }
        return options;
    }
}
