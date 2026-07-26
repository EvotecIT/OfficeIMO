namespace OfficeIMO.Html;

/// <summary>Describes one executable HTML renderer compatibility contract.</summary>
public sealed class HtmlRenderCapability {
    private readonly IReadOnlyList<string> _features;
    private readonly IReadOnlyList<string> _diagnosticCodes;

    /// <summary>Creates a renderer compatibility contract.</summary>
    public HtmlRenderCapability(
        string id,
        string area,
        HtmlRenderCapabilityKind kind,
        HtmlRenderSupportLevel supportLevel,
        IEnumerable<string> features,
        string behavior,
        IEnumerable<string>? diagnosticCodes = null) {
        Id = Required(id, nameof(id));
        Area = Required(area, nameof(area));
        Kind = kind;
        SupportLevel = supportLevel;
        Behavior = Required(behavior, nameof(behavior));
        _features = Normalize(features, nameof(features));
        _diagnosticCodes = Normalize(diagnosticCodes ?? Array.Empty<string>(), nameof(diagnosticCodes));
        if (_features.Count == 0) {
            throw new ArgumentException("At least one standards feature is required.", nameof(features));
        }
    }

    /// <summary>Stable machine-readable capability identifier.</summary>
    public string Id { get; }

    /// <summary>Human-readable renderer area.</summary>
    public string Area { get; }

    /// <summary>Standards surface represented by this entry.</summary>
    public HtmlRenderCapabilityKind Kind { get; }

    /// <summary>Observable support outcome.</summary>
    public HtmlRenderSupportLevel SupportLevel { get; }

    /// <summary>CSS properties, at-rules, elements, or artifact features covered by the entry.</summary>
    public IReadOnlyList<string> Features => _features;

    /// <summary>Exact supported subset or fallback behavior.</summary>
    public string Behavior { get; }

    /// <summary>Stable diagnostics emitted when the declared boundary is crossed.</summary>
    public IReadOnlyList<string> DiagnosticCodes => _diagnosticCodes;

    private static string Required(string value, string parameterName) {
        if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("Value cannot be empty.", parameterName);
        return value.Trim();
    }

    private static IReadOnlyList<string> Normalize(IEnumerable<string> values, string parameterName) {
        if (values == null) throw new ArgumentNullException(parameterName);
        return values
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Select(value => value.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
            .ToList()
            .AsReadOnly();
    }
}
