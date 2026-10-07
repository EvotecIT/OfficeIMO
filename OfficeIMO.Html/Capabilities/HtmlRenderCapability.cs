namespace OfficeIMO.Html;

/// <summary>Describes one executable HTML renderer compatibility contract.</summary>
public sealed class HtmlRenderCapability {
    private readonly IReadOnlyList<string> _features;
    private readonly IReadOnlyList<string> _limitations;
    private readonly IReadOnlyList<string> _diagnosticCodes;
    private readonly IReadOnlyList<HtmlCapabilityProfileBinding> _profileBindings;

    /// <summary>Creates a renderer compatibility contract.</summary>
    public HtmlRenderCapability(
        string id,
        string area,
        HtmlRenderCapabilityKind kind,
        HtmlCapabilityStage stages,
        IEnumerable<HtmlCapabilityProfileBinding> profileBindings,
        IEnumerable<string> features,
        string behavior,
        IEnumerable<string>? limitations = null,
        IEnumerable<string>? diagnosticCodes = null) {
        Id = Required(id, nameof(id));
        Area = Required(area, nameof(area));
        if (!Enum.IsDefined(typeof(HtmlRenderCapabilityKind), kind)) throw new ArgumentOutOfRangeException(nameof(kind));
        Kind = kind;
        const HtmlCapabilityStage knownStages = HtmlCapabilityStage.SourceAndDecoding
            | HtmlCapabilityStage.ParseAndPreserve
            | HtmlCapabilityStage.DomAndQuery
            | HtmlCapabilityStage.CascadeAndCompute
            | HtmlCapabilityStage.Layout
            | HtmlCapabilityStage.PaintAndOutput
            | HtmlCapabilityStage.RuntimeAndInteraction;
        if (stages == HtmlCapabilityStage.None || (stages & ~knownStages) != 0) throw new ArgumentOutOfRangeException(nameof(stages));
        Stages = stages;
        Behavior = Required(behavior, nameof(behavior));
        _profileBindings = HtmlCapabilityContractValue.Unique(profileBindings, binding => binding.ProfileId, nameof(profileBindings));
        _features = Normalize(features, nameof(features));
        _limitations = Normalize(limitations ?? Array.Empty<string>(), nameof(limitations));
        _diagnosticCodes = Normalize(diagnosticCodes ?? Array.Empty<string>(), nameof(diagnosticCodes));
        if (_profileBindings.Count == 0) {
            throw new ArgumentException("At least one compatibility profile binding is required.", nameof(profileBindings));
        }
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

    /// <summary>Processing stages covered by this claim.</summary>
    public HtmlCapabilityStage Stages { get; }

    /// <summary>Coverage, handling, maturity, promotion, provider, specification, and evidence by versioned profile.</summary>
    public IReadOnlyList<HtmlCapabilityProfileBinding> ProfileBindings => _profileBindings;

    /// <summary>CSS properties, at-rules, elements, or artifact features covered by the entry.</summary>
    public IReadOnlyList<string> Features => _features;

    /// <summary>Exact supported subset or fallback behavior.</summary>
    public string Behavior { get; }

    /// <summary>Known feature limits beyond the exact supported subset.</summary>
    public IReadOnlyList<string> Limitations => _limitations;

    /// <summary>Stable diagnostics emitted when the declared boundary is crossed.</summary>
    public IReadOnlyList<string> DiagnosticCodes => _diagnosticCodes;

    /// <summary>Gets this capability's classification for a versioned compatibility profile.</summary>
    public HtmlCapabilityProfileBinding GetProfileBinding(string profileId) {
        if (!TryGetProfileBinding(profileId, out HtmlCapabilityProfileBinding binding)) {
            throw new ArgumentOutOfRangeException(nameof(profileId), profileId, "The capability is not declared for that profile.");
        }
        return binding;
    }

    /// <summary>Attempts to get this capability's classification for a versioned compatibility profile.</summary>
    public bool TryGetProfileBinding(string? profileId, out HtmlCapabilityProfileBinding binding) {
        if (!string.IsNullOrWhiteSpace(profileId)) {
            string normalizedProfileId = profileId!.Trim();
            HtmlCapabilityProfileBinding? found = _profileBindings.FirstOrDefault(item =>
                string.Equals(item.ProfileId, normalizedProfileId, StringComparison.OrdinalIgnoreCase));
            if (found != null) {
                binding = found;
                return true;
            }
        }
        binding = null!;
        return false;
    }

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
