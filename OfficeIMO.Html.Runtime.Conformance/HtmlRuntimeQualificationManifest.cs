using System.Reflection;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace OfficeIMO.Html.Runtime.Conformance;

/// <summary>Versioned executable specification and evidence contract for a runtime profile.</summary>
public sealed class HtmlRuntimeQualificationManifest {
    /// <summary>Stable profile identity.</summary>
    public string Id { get; init; } = string.Empty;
    /// <summary>Profile contract version.</summary>
    public int Version { get; init; }
    /// <summary>Runtime profiles covered by this qualification.</summary>
    public IReadOnlyList<HtmlRuntimeProfile> RuntimeProfiles { get; init; } = Array.Empty<HtmlRuntimeProfile>();
    /// <summary>Pinned specification scopes.</summary>
    public IReadOnlyList<HtmlRuntimeSpecificationScope> Specifications { get; init; } = Array.Empty<HtmlRuntimeSpecificationScope>();
    /// <summary>Explicit upstream-suite selections and counts.</summary>
    public IReadOnlyList<HtmlRuntimeUpstreamSuiteScope> UpstreamSuites { get; init; } = Array.Empty<HtmlRuntimeUpstreamSuiteScope>();
    /// <summary>Selected executable cases.</summary>
    public IReadOnlyList<HtmlRuntimeQualificationCase> Cases { get; init; } = Array.Empty<HtmlRuntimeQualificationCase>();
    /// <summary>Expected provider evidence.</summary>
    public IReadOnlyList<HtmlRuntimeProviderExpectation> Providers { get; init; } = Array.Empty<HtmlRuntimeProviderExpectation>();
    /// <summary>Optional consuming adapter evidence.</summary>
    public IReadOnlyList<HtmlRuntimeConsumerExpectation> Consumers { get; init; } = Array.Empty<HtmlRuntimeConsumerExpectation>();
    /// <summary>Behavior explicitly excluded from the profile.</summary>
    public IReadOnlyList<string> Exclusions { get; init; } = Array.Empty<string>();
    /// <summary>Behavior still untested by this profile version.</summary>
    public IReadOnlyList<string> Untested { get; init; } = Array.Empty<string>();
}

/// <summary>One specification revision and its deliberately selected scope.</summary>
public sealed class HtmlRuntimeSpecificationScope {
    /// <summary>Stable specification identity.</summary>
    public string Id { get; init; } = string.Empty;
    /// <summary>Pinned edition or selection date.</summary>
    public string Revision { get; init; } = string.Empty;
    /// <summary>Specification location.</summary>
    public Uri Uri { get; init; } = new("https://officeimo.invalid/");
    /// <summary>Behavior claimed from this specification.</summary>
    public string Scope { get; init; } = string.Empty;
}

/// <summary>Exact selection and outcome counts for an upstream conformance suite.</summary>
public sealed class HtmlRuntimeUpstreamSuiteScope {
    /// <summary>Stable suite identity.</summary>
    public string Id { get; init; } = string.Empty;
    /// <summary>Pinned suite revision or explicit no-selection marker.</summary>
    public string Revision { get; init; } = string.Empty;
    /// <summary>Selected upstream cases.</summary>
    public int Selected { get; init; }
    /// <summary>Required selected cases.</summary>
    public int Required { get; init; }
    /// <summary>Passing selected cases.</summary>
    public int Passed { get; init; }
    /// <summary>Failing selected cases.</summary>
    public int Failed { get; init; }
    /// <summary>Explicitly excluded upstream groups.</summary>
    public int Excluded { get; init; }
    /// <summary>Declared but untested upstream groups.</summary>
    public int Untested { get; init; }
    /// <summary>Reason for the selection boundary.</summary>
    public string Note { get; init; } = string.Empty;
}

/// <summary>One named executable case and its assertion count.</summary>
public sealed class HtmlRuntimeQualificationCase {
    /// <summary>Stable conformance case identity.</summary>
    public string Id { get; init; } = string.Empty;
    /// <summary>Exact required assertion count.</summary>
    public int RequiredAssertions { get; init; }
}

/// <summary>Expected exact counts for one provider.</summary>
public sealed class HtmlRuntimeProviderExpectation {
    /// <summary>Stable provider identity.</summary>
    public string Id { get; init; } = string.Empty;
    /// <summary>Required case count.</summary>
    public int RequiredCases { get; init; }
    /// <summary>Expected passing case count.</summary>
    public int PassedCases { get; init; }
    /// <summary>Expected failing case count.</summary>
    public int FailedCases { get; init; }
    /// <summary>Required assertion count.</summary>
    public int RequiredAssertions { get; init; }
    /// <summary>Expected passing assertion count.</summary>
    public int PassedAssertions { get; init; }
    /// <summary>Expected failing assertion count.</summary>
    public int FailedAssertions { get; init; }
    /// <summary>Expected unreached assertion count.</summary>
    public int UntestedAssertions { get; init; }
}

/// <summary>Required workflow evidence for an optional consuming adapter.</summary>
public sealed class HtmlRuntimeConsumerExpectation {
    /// <summary>Stable consumer identity.</summary>
    public string Id { get; init; } = string.Empty;
    /// <summary>External contract implemented by the consumer.</summary>
    public string Contract { get; init; } = string.Empty;
    /// <summary>Pinned external package version used by the evidence.</summary>
    public string PackageVersion { get; init; } = string.Empty;
    /// <summary>Required workflow count.</summary>
    public int RequiredWorkflows { get; init; }
}

/// <summary>Actual provider counts after applying one qualification manifest.</summary>
public sealed class HtmlRuntimeQualificationResult {
    /// <summary>Qualified profile identity.</summary>
    public string ProfileId { get; init; } = string.Empty;
    /// <summary>Evaluated provider identity.</summary>
    public string ProviderId { get; init; } = string.Empty;
    /// <summary>Required case count.</summary>
    public int RequiredCases { get; init; }
    /// <summary>Passing case count.</summary>
    public int PassedCases { get; init; }
    /// <summary>Failing case count.</summary>
    public int FailedCases { get; init; }
    /// <summary>Required assertion count.</summary>
    public int RequiredAssertions { get; init; }
    /// <summary>Passing assertion count.</summary>
    public int PassedAssertions { get; init; }
    /// <summary>Failing assertion count.</summary>
    public int FailedAssertions { get; init; }
    /// <summary>Unreached assertion count.</summary>
    public int UntestedAssertions { get; init; }
    /// <summary>Whether actual counts exactly match the manifest and contain no failures.</summary>
    public bool Passed { get; init; }
}

/// <summary>Actual workflow counts for one consuming adapter.</summary>
public sealed class HtmlRuntimeConsumerQualificationResult {
    /// <summary>Qualified profile identity.</summary>
    public string ProfileId { get; init; } = string.Empty;
    /// <summary>Evaluated consumer identity.</summary>
    public string ConsumerId { get; init; } = string.Empty;
    /// <summary>Required workflow count.</summary>
    public int RequiredWorkflows { get; init; }
    /// <summary>Passing workflow count.</summary>
    public int PassedWorkflows { get; init; }
    /// <summary>Failing workflow count.</summary>
    public int FailedWorkflows { get; init; }
    /// <summary>Whether actual counts satisfy the consumer requirement without failures.</summary>
    public bool Passed { get; init; }
}

/// <summary>Combined provider and consumer evidence for one qualification profile.</summary>
public sealed class HtmlRuntimeProfileQualificationResult {
    /// <summary>Qualified profile identity.</summary>
    public string ProfileId { get; init; } = string.Empty;
    /// <summary>Provider qualification evidence.</summary>
    public HtmlRuntimeQualificationResult Provider { get; init; } = null!;
    /// <summary>Consumer workflow evidence.</summary>
    public IReadOnlyList<HtmlRuntimeConsumerQualificationResult> Consumers { get; init; } = Array.Empty<HtmlRuntimeConsumerQualificationResult>();
    /// <summary>Whether the provider and every declared consumer passed.</summary>
    public bool Passed { get; init; }
}

/// <summary>Loads and validates the embedded v1 profile manifests.</summary>
public static class HtmlRuntimeQualificationCatalog {
    private static readonly Lazy<IReadOnlyList<HtmlRuntimeQualificationManifest>> Manifests = new(Load);
    /// <summary>All validated embedded manifests.</summary>
    public static IReadOnlyList<HtmlRuntimeQualificationManifest> All => Manifests.Value;
    /// <summary>Returns one manifest by stable identity.</summary>
    public static HtmlRuntimeQualificationManifest Get(string id) => All.Single(item => string.Equals(item.Id, id, StringComparison.Ordinal));

    /// <summary>Evaluates a provider report against the exact selected cases and expected counts.</summary>
    public static HtmlRuntimeQualificationResult EvaluateProvider(HtmlRuntimeQualificationManifest manifest, HtmlRuntimeConformanceReport report) {
        ArgumentNullException.ThrowIfNull(manifest);
        ArgumentNullException.ThrowIfNull(report);
        HtmlRuntimeProviderExpectation expected = manifest.Providers.Single(item => item.Id == report.Provider.Id);
        var selected = manifest.Cases.Select(item => report.Cases.Single(result => result.Id == item.Id)).ToArray();
        var actual = new HtmlRuntimeQualificationResult {
            ProfileId = manifest.Id, ProviderId = report.Provider.Id,
            RequiredCases = selected.Length,
            PassedCases = selected.Count(item => item.Passed), FailedCases = selected.Count(item => !item.Passed),
            RequiredAssertions = selected.Sum(item => item.RequiredAssertions),
            PassedAssertions = selected.Sum(item => item.PassedAssertions),
            FailedAssertions = selected.Sum(item => item.FailedAssertions),
            UntestedAssertions = selected.Sum(item => item.UntestedAssertions)
        };
        bool matches = actual.RequiredCases == expected.RequiredCases && actual.PassedCases == expected.PassedCases
            && actual.FailedCases == expected.FailedCases && actual.RequiredAssertions == expected.RequiredAssertions
            && actual.PassedAssertions == expected.PassedAssertions && actual.FailedAssertions == expected.FailedAssertions
            && actual.UntestedAssertions == expected.UntestedAssertions;
        return new HtmlRuntimeQualificationResult {
            ProfileId = actual.ProfileId, ProviderId = actual.ProviderId,
            RequiredCases = actual.RequiredCases, PassedCases = actual.PassedCases, FailedCases = actual.FailedCases,
            RequiredAssertions = actual.RequiredAssertions, PassedAssertions = actual.PassedAssertions,
            FailedAssertions = actual.FailedAssertions, UntestedAssertions = actual.UntestedAssertions,
            Passed = matches && actual.FailedCases == 0 && actual.FailedAssertions == 0 && actual.UntestedAssertions == 0
        };
    }

    /// <summary>Evaluates actual workflow outcomes for one declared consumer.</summary>
    public static HtmlRuntimeConsumerQualificationResult EvaluateConsumer(HtmlRuntimeQualificationManifest manifest,
        string consumerId, int passedWorkflows, int failedWorkflows) {
        ArgumentNullException.ThrowIfNull(manifest);
        ArgumentException.ThrowIfNullOrWhiteSpace(consumerId);
        if (passedWorkflows < 0) throw new ArgumentOutOfRangeException(nameof(passedWorkflows));
        if (failedWorkflows < 0) throw new ArgumentOutOfRangeException(nameof(failedWorkflows));
        HtmlRuntimeConsumerExpectation expected = manifest.Consumers.Single(item => item.Id == consumerId);
        return new HtmlRuntimeConsumerQualificationResult {
            ProfileId = manifest.Id,
            ConsumerId = consumerId,
            RequiredWorkflows = expected.RequiredWorkflows,
            PassedWorkflows = passedWorkflows,
            FailedWorkflows = failedWorkflows,
            Passed = passedWorkflows == expected.RequiredWorkflows && failedWorkflows == 0
        };
    }

    /// <summary>Combines provider evidence with every consumer workflow declared by the profile.</summary>
    public static HtmlRuntimeProfileQualificationResult Evaluate(HtmlRuntimeQualificationManifest manifest,
        HtmlRuntimeConformanceReport report, params HtmlRuntimeConsumerQualificationResult[] consumers) {
        ArgumentNullException.ThrowIfNull(manifest);
        ArgumentNullException.ThrowIfNull(consumers);
        HtmlRuntimeQualificationResult provider = EvaluateProvider(manifest, report);
        HtmlRuntimeConsumerQualificationResult[] actual = consumers.ToArray();
        bool consumerEvidenceMatches = actual.Length == manifest.Consumers.Count
            && actual.Select(item => item.ConsumerId).Distinct(StringComparer.Ordinal).Count() == actual.Length
            && manifest.Consumers.All(expected => actual.Any(item =>
                item.ProfileId == manifest.Id && item.ConsumerId == expected.Id
                && item.RequiredWorkflows == expected.RequiredWorkflows
                && item.PassedWorkflows == expected.RequiredWorkflows
                && item.FailedWorkflows == 0 && item.Passed));
        return new HtmlRuntimeProfileQualificationResult {
            ProfileId = manifest.Id,
            Provider = provider,
            Consumers = Array.AsReadOnly(actual),
            Passed = provider.Passed && consumerEvidenceMatches
        };
    }

    private static IReadOnlyList<HtmlRuntimeQualificationManifest> Load() {
        using Stream stream = typeof(HtmlRuntimeQualificationCatalog).Assembly.GetManifestResourceStream("OfficeIMO.Html.Runtime.Conformance.Manifests.runtime-profiles-v1.json")
            ?? throw new InvalidOperationException("The runtime qualification manifest is missing.");
        HtmlRuntimeQualificationDocument document = JsonSerializer.Deserialize(stream, HtmlRuntimeQualificationJsonContext.Default.HtmlRuntimeQualificationDocument)
            ?? throw new InvalidOperationException("The runtime qualification manifest is empty.");
        if (document.SchemaVersion != 1 || document.SuiteId != "officeimo-html-runtime-conformance-v1" || document.Profiles.Count != 3)
            throw new InvalidOperationException("The runtime qualification manifest header is invalid.");
        foreach (HtmlRuntimeQualificationManifest manifest in document.Profiles) Validate(manifest);
        return Array.AsReadOnly(document.Profiles.ToArray());
    }

    private static void Validate(HtmlRuntimeQualificationManifest manifest) {
        if (string.IsNullOrWhiteSpace(manifest.Id) || manifest.Version != 1 || manifest.Specifications.Count == 0
            || manifest.Cases.Count == 0 || manifest.Providers.Count == 0 || manifest.Exclusions.Count == 0 || manifest.Untested.Count == 0)
            throw new InvalidOperationException($"Qualification profile '{manifest.Id}' is incomplete.");
        if (manifest.Cases.Select(item => item.Id).Distinct(StringComparer.Ordinal).Count() != manifest.Cases.Count
            || manifest.Providers.Select(item => item.Id).Distinct(StringComparer.Ordinal).Count() != manifest.Providers.Count
            || manifest.Consumers.Select(item => item.Id).Distinct(StringComparer.Ordinal).Count() != manifest.Consumers.Count)
            throw new InvalidOperationException($"Qualification profile '{manifest.Id}' contains duplicate identities.");
        foreach (HtmlRuntimeQualificationCase item in manifest.Cases) {
            if (!HtmlRuntimeConformanceSuite.CaseAssertions.TryGetValue(item.Id, out int assertions) || assertions != item.RequiredAssertions)
                throw new InvalidOperationException($"Qualification case '{item.Id}' does not match the executable suite.");
        }
        int caseCount = manifest.Cases.Count;
        int assertionCount = manifest.Cases.Sum(item => item.RequiredAssertions);
        foreach (HtmlRuntimeProviderExpectation provider in manifest.Providers) {
            if (provider.RequiredCases != caseCount || provider.RequiredAssertions != assertionCount
                || provider.PassedCases + provider.FailedCases != provider.RequiredCases
                || provider.PassedAssertions + provider.FailedAssertions + provider.UntestedAssertions != provider.RequiredAssertions)
                throw new InvalidOperationException($"Qualification counts for provider '{provider.Id}' are inconsistent.");
        }
        foreach (HtmlRuntimeConsumerExpectation consumer in manifest.Consumers ?? Array.Empty<HtmlRuntimeConsumerExpectation>()) {
            if (string.IsNullOrWhiteSpace(consumer.Id) || string.IsNullOrWhiteSpace(consumer.Contract)
                || string.IsNullOrWhiteSpace(consumer.PackageVersion) || consumer.RequiredWorkflows <= 0)
                throw new InvalidOperationException($"Consumer evidence for '{consumer.Id}' is inconsistent.");
        }
        foreach (HtmlRuntimeUpstreamSuiteScope suite in manifest.UpstreamSuites ?? Array.Empty<HtmlRuntimeUpstreamSuiteScope>()) {
            if (suite.Selected != suite.Required || suite.Passed + suite.Failed != suite.Required
                || suite.Required < 0 || suite.Excluded < 0 || suite.Untested < 0)
                throw new InvalidOperationException($"Upstream suite counts for '{suite.Id}' are inconsistent.");
        }
    }
}

internal sealed class HtmlRuntimeQualificationDocument {
    public int SchemaVersion { get; init; }
    public string SuiteId { get; init; } = string.Empty;
    public IReadOnlyList<HtmlRuntimeQualificationManifest> Profiles { get; init; } = Array.Empty<HtmlRuntimeQualificationManifest>();
}

[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase, UseStringEnumConverter = true)]
[JsonSerializable(typeof(HtmlRuntimeQualificationDocument))]
[JsonSerializable(typeof(HtmlRuntimeQualificationManifest))]
[JsonSerializable(typeof(HtmlRuntimeQualificationResult))]
[JsonSerializable(typeof(HtmlRuntimeConsumerQualificationResult))]
[JsonSerializable(typeof(HtmlRuntimeProfileQualificationResult))]
internal sealed partial class HtmlRuntimeQualificationJsonContext : JsonSerializerContext;
