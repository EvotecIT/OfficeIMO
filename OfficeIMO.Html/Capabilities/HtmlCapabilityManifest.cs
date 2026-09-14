namespace OfficeIMO.Html;

/// <summary>Identifies one implementation provider pinned by a compatibility profile.</summary>
public sealed class HtmlCapabilityProviderPin {
    /// <summary>Creates a provider pin.</summary>
    public HtmlCapabilityProviderPin(string id, string name, string version, HtmlCapabilityProviderOwnership ownership) {
        Id = HtmlCapabilityContractValue.Required(id, nameof(id));
        Name = HtmlCapabilityContractValue.Required(name, nameof(name));
        Version = HtmlCapabilityContractValue.Required(version, nameof(version));
        if (!Enum.IsDefined(typeof(HtmlCapabilityProviderOwnership), ownership)) throw new ArgumentOutOfRangeException(nameof(ownership));
        Ownership = ownership;
    }

    /// <summary>Stable provider identifier.</summary>
    public string Id { get; }
    /// <summary>Human-readable provider name.</summary>
    public string Name { get; }
    /// <summary>Version or version-resolution contract pinned by the profile.</summary>
    public string Version { get; }
    /// <summary>Provider ownership classification.</summary>
    public HtmlCapabilityProviderOwnership Ownership { get; }
}

/// <summary>Identifies an immutable specification source selected by a compatibility profile.</summary>
public sealed class HtmlCapabilitySpecificationPin {
    /// <summary>Creates a specification pin.</summary>
    public HtmlCapabilitySpecificationPin(string id, string title, string uri, string revision, string scope) {
        Id = HtmlCapabilityContractValue.Required(id, nameof(id));
        Title = HtmlCapabilityContractValue.Required(title, nameof(title));
        Uri = HtmlCapabilityContractValue.Required(uri, nameof(uri));
        Revision = HtmlCapabilityContractValue.Required(revision, nameof(revision));
        Scope = HtmlCapabilityContractValue.Required(scope, nameof(scope));
    }

    /// <summary>Stable specification identifier used by capability bindings.</summary>
    public string Id { get; }
    /// <summary>Human-readable specification title.</summary>
    public string Title { get; }
    /// <summary>Immutable snapshot or revision-addressed source URI.</summary>
    public string Uri { get; }
    /// <summary>Exact published revision, repository commit, or data version.</summary>
    public string Revision { get; }
    /// <summary>Sections or module scope selected from the pinned source.</summary>
    public string Scope { get; }
}

/// <summary>Identifies a versioned evidence source selected by a compatibility profile.</summary>
public sealed class HtmlCapabilityEvidencePin {
    private readonly IReadOnlyList<string> _caseIds;
    private readonly IReadOnlyList<HtmlCapabilityEvidenceSelection> _selections;

    /// <summary>Creates an evidence pin.</summary>
    public HtmlCapabilityEvidencePin(
        string id,
        string source,
        string revision,
        HtmlCapabilityEvidenceRole role,
        string scope,
        int? required = null,
        int? passed = null,
        int? failed = null,
        int? excluded = null,
        int? untested = null,
        IEnumerable<string>? caseIds = null,
        IEnumerable<HtmlCapabilityEvidenceSelection>? selections = null) {
        Id = HtmlCapabilityContractValue.Required(id, nameof(id));
        Source = HtmlCapabilityContractValue.Required(source, nameof(source));
        Revision = HtmlCapabilityContractValue.Required(revision, nameof(revision));
        if (!Enum.IsDefined(typeof(HtmlCapabilityEvidenceRole), role)) throw new ArgumentOutOfRangeException(nameof(role));
        Role = role;
        Scope = HtmlCapabilityContractValue.Required(scope, nameof(scope));
        Required = NonNegative(required, nameof(required));
        Passed = NonNegative(passed, nameof(passed));
        Failed = NonNegative(failed, nameof(failed));
        Excluded = NonNegative(excluded, nameof(excluded));
        Untested = NonNegative(untested, nameof(untested));
        _caseIds = HtmlCapabilityContractValue.Normalize(caseIds ?? Array.Empty<string>(), nameof(caseIds));
        _selections = HtmlCapabilityContractValue.Unique(
            selections ?? Array.Empty<HtmlCapabilityEvidenceSelection>(),
            item => item.CapabilityId,
            nameof(selections));
        if (Required.HasValue && Passed.HasValue && Failed.HasValue && Passed.Value + Failed.Value > Required.Value) {
            throw new ArgumentException("Passed and failed evidence counts cannot exceed the required count.");
        }
    }

    /// <summary>Stable evidence identifier used by capability bindings.</summary>
    public string Id { get; }
    /// <summary>Test suite, corpus, or comparison source.</summary>
    public string Source { get; }
    /// <summary>Immutable source commit, manifest version, or package-source contract.</summary>
    public string Revision { get; }
    /// <summary>How the evidence contributes to compatibility qualification.</summary>
    public HtmlCapabilityEvidenceRole Role { get; }
    /// <summary>Included and excluded evidence scope.</summary>
    public string Scope { get; }
    /// <summary>Number of cases required by the selected manifest, when recorded.</summary>
    public int? Required { get; }
    /// <summary>Number of required cases that passed, when recorded.</summary>
    public int? Passed { get; }
    /// <summary>Number of required cases that failed, when recorded.</summary>
    public int? Failed { get; }
    /// <summary>Number of cases explicitly excluded from the selected manifest, when recorded.</summary>
    public int? Excluded { get; }
    /// <summary>Number of selected cases not yet executed, when recorded.</summary>
    public int? Untested { get; }
    /// <summary>Stable identifiers for the selected required cases, when the evidence is count-based.</summary>
    public IReadOnlyList<string> CaseIds => _caseIds;
    /// <summary>Exact per-capability case selections and feature exclusions.</summary>
    public IReadOnlyList<HtmlCapabilityEvidenceSelection> Selections => _selections;

    private static int? NonNegative(int? value, string parameterName) {
        if (value < 0) throw new ArgumentOutOfRangeException(parameterName);
        return value;
    }
}

/// <summary>Defines one versioned compatibility profile and its immutable provider, specification, and evidence pins.</summary>
public sealed class HtmlCapabilityProfileManifest {
    private readonly IReadOnlyList<HtmlCapabilityProviderPin> _providers;
    private readonly IReadOnlyList<HtmlCapabilitySpecificationPin> _specifications;
    private readonly IReadOnlyList<HtmlCapabilityEvidencePin> _evidence;
    private readonly IReadOnlyList<string> _platforms;
    private readonly IReadOnlyList<string> _outputs;

    /// <summary>Creates a compatibility profile manifest.</summary>
    public HtmlCapabilityProfileManifest(
        string id,
        string version,
        string title,
        HtmlCapabilityPromotionState promotion,
        IEnumerable<HtmlCapabilityProviderPin> providers,
        IEnumerable<HtmlCapabilitySpecificationPin> specifications,
        IEnumerable<HtmlCapabilityEvidencePin> evidence,
        IEnumerable<string> platforms,
        IEnumerable<string> outputs) {
        Id = HtmlCapabilityContractValue.Required(id, nameof(id));
        Version = HtmlCapabilityContractValue.Required(version, nameof(version));
        Title = HtmlCapabilityContractValue.Required(title, nameof(title));
        if (!Enum.IsDefined(typeof(HtmlCapabilityPromotionState), promotion)) throw new ArgumentOutOfRangeException(nameof(promotion));
        Promotion = promotion;
        _providers = HtmlCapabilityContractValue.Unique(providers, item => item.Id, nameof(providers));
        _specifications = HtmlCapabilityContractValue.Unique(specifications, item => item.Id, nameof(specifications));
        _evidence = HtmlCapabilityContractValue.Unique(evidence, item => item.Id, nameof(evidence));
        _platforms = HtmlCapabilityContractValue.Normalize(platforms, nameof(platforms));
        _outputs = HtmlCapabilityContractValue.Normalize(outputs, nameof(outputs));
        if (_providers.Count == 0 || _specifications.Count == 0 || _evidence.Count == 0 || _platforms.Count == 0 || _outputs.Count == 0) {
            throw new ArgumentException("A compatibility profile requires providers, specifications, evidence, platforms, and outputs.");
        }
    }

    /// <summary>Stable versioned profile identifier.</summary>
    public string Id { get; }
    /// <summary>Profile contract version.</summary>
    public string Version { get; }
    /// <summary>Human-readable profile title.</summary>
    public string Title { get; }
    /// <summary>Current release-promotion state.</summary>
    public HtmlCapabilityPromotionState Promotion { get; }
    /// <summary>Providers admitted by the profile.</summary>
    public IReadOnlyList<HtmlCapabilityProviderPin> Providers => _providers;
    /// <summary>Immutable specification sources admitted by the profile.</summary>
    public IReadOnlyList<HtmlCapabilitySpecificationPin> Specifications => _specifications;
    /// <summary>Versioned qualification, regression, and reference evidence.</summary>
    public IReadOnlyList<HtmlCapabilityEvidencePin> Evidence => _evidence;
    /// <summary>Platforms covered by this manifest.</summary>
    public IReadOnlyList<string> Platforms => _platforms;
    /// <summary>Output contracts covered by this manifest.</summary>
    public IReadOnlyList<string> Outputs => _outputs;
}

/// <summary>Classifies one capability within one versioned compatibility profile.</summary>
public sealed class HtmlCapabilityProfileBinding {
    private readonly IReadOnlyList<string> _providerIds;
    private readonly IReadOnlyList<string> _optionalProviderIds;
    private readonly IReadOnlyList<string> _specificationIds;
    private readonly IReadOnlyList<string> _evidenceIds;

    /// <summary>Creates a capability-to-profile binding.</summary>
    public HtmlCapabilityProfileBinding(
        string profileId,
        HtmlCapabilityCoverage coverage,
        HtmlCapabilityHandling handling,
        HtmlCapabilityMaturity maturity,
        HtmlCapabilityPromotionState promotion,
        IEnumerable<string> providerIds,
        IEnumerable<string> specificationIds,
        IEnumerable<string> evidenceIds,
        IEnumerable<string>? optionalProviderIds = null) {
        ProfileId = HtmlCapabilityContractValue.Required(profileId, nameof(profileId));
        if (!Enum.IsDefined(typeof(HtmlCapabilityCoverage), coverage)) throw new ArgumentOutOfRangeException(nameof(coverage));
        if (!Enum.IsDefined(typeof(HtmlCapabilityHandling), handling)) throw new ArgumentOutOfRangeException(nameof(handling));
        if (!Enum.IsDefined(typeof(HtmlCapabilityMaturity), maturity)) throw new ArgumentOutOfRangeException(nameof(maturity));
        if (!Enum.IsDefined(typeof(HtmlCapabilityPromotionState), promotion)) throw new ArgumentOutOfRangeException(nameof(promotion));
        Coverage = coverage;
        Handling = handling;
        Maturity = maturity;
        Promotion = promotion;
        _providerIds = HtmlCapabilityContractValue.Normalize(providerIds, nameof(providerIds));
        _optionalProviderIds = HtmlCapabilityContractValue.Normalize(optionalProviderIds ?? Array.Empty<string>(), nameof(optionalProviderIds));
        _specificationIds = HtmlCapabilityContractValue.Normalize(specificationIds, nameof(specificationIds));
        _evidenceIds = HtmlCapabilityContractValue.Normalize(evidenceIds, nameof(evidenceIds));
        if (_providerIds.Count == 0 || _specificationIds.Count == 0 || _evidenceIds.Count == 0) {
            throw new ArgumentException("A profile binding requires providers, specifications, and evidence.");
        }
    }

    /// <summary>Profile manifest identifier.</summary>
    public string ProfileId { get; }
    /// <summary>Qualified implementation coverage for this profile.</summary>
    public HtmlCapabilityCoverage Coverage { get; }
    /// <summary>Observable handling for this profile.</summary>
    public HtmlCapabilityHandling Handling { get; }
    /// <summary>Requirement maturity for this profile.</summary>
    public HtmlCapabilityMaturity Maturity { get; }
    /// <summary>Release-promotion state for this profile.</summary>
    public HtmlCapabilityPromotionState Promotion { get; }
    /// <summary>Provider identifiers required by this capability.</summary>
    public IReadOnlyList<string> ProviderIds => _providerIds;
    /// <summary>Provider identifiers used only when explicitly supplied or selected.</summary>
    public IReadOnlyList<string> OptionalProviderIds => _optionalProviderIds;
    /// <summary>Specification identifiers defining the supported subset.</summary>
    public IReadOnlyList<string> SpecificationIds => _specificationIds;
    /// <summary>Evidence identifiers supporting the classification.</summary>
    public IReadOnlyList<string> EvidenceIds => _evidenceIds;
}

internal static class HtmlCapabilityContractValue {
    internal static string Required(string value, string parameterName) {
        if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("Value cannot be empty.", parameterName);
        return value.Trim();
    }

    internal static IReadOnlyList<string> Normalize(IEnumerable<string> values, string parameterName) {
        if (values == null) throw new ArgumentNullException(parameterName);
        return values.Where(value => !string.IsNullOrWhiteSpace(value))
            .Select(value => value.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
            .ToList()
            .AsReadOnly();
    }

    internal static IReadOnlyList<T> Unique<T>(IEnumerable<T> values, Func<T, string> id, string parameterName) {
        if (values == null) throw new ArgumentNullException(parameterName);
        T[] result = values.ToArray();
        if (result.Any(item => item == null)) throw new ArgumentException("Entries cannot be null.", parameterName);
        string[] duplicates = result.GroupBy(id, StringComparer.OrdinalIgnoreCase)
            .Where(group => group.Count() > 1)
            .Select(group => group.Key)
            .ToArray();
        if (duplicates.Length != 0) {
            throw new ArgumentException("Identifiers must be unique: " + string.Join(", ", duplicates) + ".", parameterName);
        }
        return Array.AsReadOnly(result);
    }
}
