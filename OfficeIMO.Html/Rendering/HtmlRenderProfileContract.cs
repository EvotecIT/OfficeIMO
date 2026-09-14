namespace OfficeIMO.Html;

/// <summary>Executable defaults and qualification state for one named HTML render intent.</summary>
public sealed class HtmlRenderProfileContract {
    private readonly IReadOnlyList<string> _capabilityProfileIds;
    private readonly IReadOnlyList<HtmlRenderEncoder> _encoders;
    private readonly IReadOnlyList<HtmlRenderPageSetMode> _pageSets;
    private readonly IReadOnlyList<string> _evidenceIds;

    internal HtmlRenderProfileContract(
        HtmlRenderIntentProfile profile,
        string id,
        string name,
        IEnumerable<string> capabilityProfileIds,
        HtmlCssMediaContext cssMedia,
        HtmlRenderLayoutSurface surface,
        HtmlRenderPaginationPolicy pagination,
        HtmlRenderPageSet defaultPageSet,
        HtmlCapabilityCoverage coverage,
        HtmlCapabilityPromotionState promotion,
        IEnumerable<HtmlRenderEncoder> encoders,
        IEnumerable<HtmlRenderPageSetMode> pageSets,
        IEnumerable<string> evidenceIds,
        string behavior,
        string limitations) {
        Profile = profile;
        Id = Required(id, nameof(id));
        Name = Required(name, nameof(name));
        _capabilityProfileIds = Snapshot(capabilityProfileIds, nameof(capabilityProfileIds));
        if (_capabilityProfileIds.Count == 0) {
            throw new ArgumentException("At least one compatibility profile is required.", nameof(capabilityProfileIds));
        }
        CssMedia = cssMedia;
        Surface = surface;
        Pagination = pagination;
        DefaultPageSet = defaultPageSet ?? throw new ArgumentNullException(nameof(defaultPageSet));
        Coverage = coverage;
        Promotion = promotion;
        _encoders = Snapshot(encoders, nameof(encoders));
        _pageSets = Snapshot(pageSets, nameof(pageSets));
        _evidenceIds = Snapshot(evidenceIds, nameof(evidenceIds));
        Behavior = Required(behavior, nameof(behavior));
        Limitations = limitations?.Trim() ?? string.Empty;
    }

    /// <summary>Typed profile identifier.</summary>
    public HtmlRenderIntentProfile Profile { get; }
    /// <summary>Stable versioned identifier.</summary>
    public string Id { get; }
    /// <summary>Human-readable profile name.</summary>
    public string Name { get; }
    /// <summary>Compatibility manifests that own the underlying layout and output evidence.</summary>
    public IReadOnlyList<string> CapabilityProfileIds => _capabilityProfileIds;
    /// <summary>CSS media type selected independently from pagination.</summary>
    public HtmlCssMediaContext CssMedia { get; }
    /// <summary>Default layout surface.</summary>
    public HtmlRenderLayoutSurface Surface { get; }
    /// <summary>Default pagination policy.</summary>
    public HtmlRenderPaginationPolicy Pagination { get; }
    /// <summary>Default explicit page-set behavior.</summary>
    public HtmlRenderPageSet DefaultPageSet { get; }
    /// <summary>Qualification of this exact axis combination.</summary>
    public HtmlCapabilityCoverage Coverage { get; }
    /// <summary>Release exposure of this exact axis combination.</summary>
    public HtmlCapabilityPromotionState Promotion { get; }
    /// <summary>Output adapters admitted for this profile.</summary>
    public IReadOnlyList<HtmlRenderEncoder> Encoders => _encoders;
    /// <summary>Page-set modes admitted for this profile.</summary>
    public IReadOnlyList<HtmlRenderPageSetMode> PageSets => _pageSets;
    /// <summary>Evidence pins supporting the qualification claim.</summary>
    public IReadOnlyList<string> EvidenceIds => _evidenceIds;
    /// <summary>Observable layout behavior.</summary>
    public string Behavior { get; }
    /// <summary>Known boundaries for this profile.</summary>
    public string Limitations { get; }

    private static string Required(string value, string parameterName) {
        if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("Value cannot be empty.", parameterName);
        return value.Trim();
    }

    private static IReadOnlyList<T> Snapshot<T>(IEnumerable<T> values, string parameterName) {
        if (values == null) throw new ArgumentNullException(parameterName);
        return values.Distinct().ToList().AsReadOnly();
    }
}
