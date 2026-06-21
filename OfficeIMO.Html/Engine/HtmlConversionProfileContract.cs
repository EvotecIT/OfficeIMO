namespace OfficeIMO.Html;

/// <summary>
/// Describes the behavior callers can expect from a shared OfficeIMO HTML conversion profile.
/// </summary>
public sealed class HtmlConversionProfileContract {
    /// <summary>
    /// Creates a conversion profile contract.
    /// </summary>
    public HtmlConversionProfileContract(
        HtmlConversionProfile profile,
        string name,
        string intendedUse,
        string fidelityGoal,
        IEnumerable<string> supportedHtml,
        IEnumerable<string> supportedCss,
        IEnumerable<string> resourceGuarantees,
        IEnumerable<string> diagnosticGuarantees) {
        Profile = profile;
        Name = name ?? throw new ArgumentNullException(nameof(name));
        IntendedUse = intendedUse ?? throw new ArgumentNullException(nameof(intendedUse));
        FidelityGoal = fidelityGoal ?? throw new ArgumentNullException(nameof(fidelityGoal));
        SupportedHtml = ToReadOnlyList(supportedHtml, nameof(supportedHtml));
        SupportedCss = ToReadOnlyList(supportedCss, nameof(supportedCss));
        ResourceGuarantees = ToReadOnlyList(resourceGuarantees, nameof(resourceGuarantees));
        DiagnosticGuarantees = ToReadOnlyList(diagnosticGuarantees, nameof(diagnosticGuarantees));
    }

    /// <summary>Profile identifier used by adapters and gallery manifests.</summary>
    public HtmlConversionProfile Profile { get; }

    /// <summary>Human readable profile name.</summary>
    public string Name { get; }

    /// <summary>Primary scenario the profile is optimized for.</summary>
    public string IntendedUse { get; }

    /// <summary>Fidelity target expressed as a caller-facing contract.</summary>
    public string FidelityGoal { get; }

    /// <summary>HTML structures the profile treats as supported capability, not accidental behavior.</summary>
    public IReadOnlyList<string> SupportedHtml { get; }

    /// <summary>CSS capabilities the profile treats as supported capability, not accidental behavior.</summary>
    public IReadOnlyList<string> SupportedCss { get; }

    /// <summary>Resource handling guarantees such as policy enforcement and manifest reporting.</summary>
    public IReadOnlyList<string> ResourceGuarantees { get; }

    /// <summary>Diagnostics callers can rely on when content is simplified, rejected, or degraded.</summary>
    public IReadOnlyList<string> DiagnosticGuarantees { get; }

    private static IReadOnlyList<string> ToReadOnlyList(IEnumerable<string> values, string parameterName) {
        if (values == null) {
            throw new ArgumentNullException(parameterName);
        }

        return values.Where(value => !string.IsNullOrWhiteSpace(value))
            .Select(value => value.Trim())
            .ToList()
            .AsReadOnly();
    }
}
