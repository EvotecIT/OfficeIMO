namespace OfficeIMO.Html;

/// <summary>Describes one built-in HTML target, its public APIs, and feature-level contract.</summary>
public sealed class HtmlTargetCapabilityContract {
    /// <summary>Creates a target contract from independently classified conversion routes.</summary>
    public HtmlTargetCapabilityContract(
        HtmlConversionTarget target,
        string packageName,
        string artifactName,
        HtmlToTargetCapabilityContract htmlToTarget,
        TargetToHtmlCapabilityContract? targetToHtml = null) {
        Target = target;
        PackageName = Required(packageName, nameof(packageName));
        ArtifactName = Required(artifactName, nameof(artifactName));
        HtmlToTarget = htmlToTarget ?? throw new ArgumentNullException(nameof(htmlToTarget));
        TargetToHtml = targetToHtml;
    }

    /// <summary>Target identifier used by preflight and adapter selection.</summary>
    public HtmlConversionTarget Target { get; }
    /// <summary>Package that owns the thin target projection.</summary>
    public string PackageName { get; }
    /// <summary>Native or rendered artifact produced by the target.</summary>
    public string ArtifactName { get; }
    /// <summary>HTML-to-target route, including its own profiles and feature partition.</summary>
    public HtmlToTargetCapabilityContract HtmlToTarget { get; }
    /// <summary>Target-to-HTML route, or <see langword="null"/> when unavailable.</summary>
    public TargetToHtmlCapabilityContract? TargetToHtml { get; }
    /// <summary>Primary public HTML import entry point.</summary>
    public string ImportEntryPoint => HtmlToTarget.EntryPoint;
    /// <summary>Public result or evidence contract returned by the import path.</summary>
    public string ImportResultContract => HtmlToTarget.ResultContract;
    /// <summary>Primary reverse HTML entry point, or <see langword="null"/> when no reverse path exists.</summary>
    public string? ExportEntryPoint => TargetToHtml?.EntryPoint;
    /// <summary>Reverse conversion evidence contract, or <see langword="null"/> when unavailable.</summary>
    public string? ExportResultContract => TargetToHtml?.ResultContract;
    /// <summary>Whether the target exposes a reverse artifact-to-HTML route.</summary>
    public bool SupportsReverseHtml => TargetToHtml != null;

    private static string Required(string value, string parameterName) {
        if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("Value cannot be empty.", parameterName);
        return value.Trim();
    }

}
