namespace OfficeIMO.Html;

/// <summary>
/// Describes a source-specific Office-to-HTML lane and how it maps to the shared HTML profile vocabulary.
/// </summary>
public sealed class OfficeHtmlConversionProfileContract {
    /// <summary>
    /// Creates a source-specific Office-to-HTML profile contract.
    /// </summary>
    public OfficeHtmlConversionProfileContract(
        OfficeHtmlConversionProfile profile,
        string sourceFormat,
        string name,
        HtmlConversionProfile sharedProfile,
        string intendedUse,
        string fidelityGoal,
        string visualPrimitiveOwner,
        IEnumerable<string> supportedHtml,
        IEnumerable<string> resourceGuarantees,
        IEnumerable<string> diagnosticGuarantees) {
        Profile = profile;
        SourceFormat = sourceFormat ?? throw new ArgumentNullException(nameof(sourceFormat));
        Name = name ?? throw new ArgumentNullException(nameof(name));
        SharedProfile = sharedProfile;
        IntendedUse = intendedUse ?? throw new ArgumentNullException(nameof(intendedUse));
        FidelityGoal = fidelityGoal ?? throw new ArgumentNullException(nameof(fidelityGoal));
        VisualPrimitiveOwner = visualPrimitiveOwner ?? throw new ArgumentNullException(nameof(visualPrimitiveOwner));
        SupportedHtml = ToReadOnlyList(supportedHtml, nameof(supportedHtml));
        ResourceGuarantees = ToReadOnlyList(resourceGuarantees, nameof(resourceGuarantees));
        DiagnosticGuarantees = ToReadOnlyList(diagnosticGuarantees, nameof(diagnosticGuarantees));
    }

    /// <summary>Source-specific profile identifier.</summary>
    public OfficeHtmlConversionProfile Profile { get; }

    /// <summary>Office source format that owns source semantics.</summary>
    public string SourceFormat { get; }

    /// <summary>Human-readable profile name.</summary>
    public string Name { get; }

    /// <summary>Shared HTML profile this source-specific profile reports against.</summary>
    public HtmlConversionProfile SharedProfile { get; }

    /// <summary>Primary scenario the lane is optimized for.</summary>
    public string IntendedUse { get; }

    /// <summary>Fidelity target expressed as a caller-facing contract.</summary>
    public string FidelityGoal { get; }

    /// <summary>Shared owner for reusable visual primitives, or "none" for semantic-only lanes.</summary>
    public string VisualPrimitiveOwner { get; }

    /// <summary>HTML structures the lane treats as supported capability.</summary>
    public IReadOnlyList<string> SupportedHtml { get; }

    /// <summary>Resource handling guarantees such as asset inventory and media policy.</summary>
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
