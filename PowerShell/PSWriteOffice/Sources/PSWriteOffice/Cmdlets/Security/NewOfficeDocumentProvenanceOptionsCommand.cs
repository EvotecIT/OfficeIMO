using System;
using System.Management.Automation;
using OfficeIMO.Provenance;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Security;

/// <summary>Creates structural, Unicode integrity and optional content-credential verification policy.</summary>
/// <example><summary>Inspect credentials without scanning text.</summary>
/// <code>$policy = New-OfficeDocumentProvenanceOptions -NoTextIntegrity -VerificationTimeoutSeconds 60
/// Get-OfficeDocumentProvenance -Path ./Press/*.png -Options $policy</code></example>
[Cmdlet(VerbsCommon.New, "OfficeDocumentProvenanceOptions")]
[OutputType(typeof(OfficeProvenanceAssessmentOptions))]
public sealed class NewOfficeDocumentProvenanceOptionsCommand : PSCmdlet {
    /// <summary>Skip Unicode text-integrity inspection.</summary>
    [Parameter]
    public SwitchParameter NoTextIntegrity { get; set; }
    /// <summary>Skip images and media embedded in containers.</summary>
    [Parameter]
    public SwitchParameter NoEmbeddedAssets { get; set; }
    /// <summary>Maximum source asset bytes.</summary>
    [Parameter]
    [ValidateRange(1, long.MaxValue)]
    public long MaximumAssetBytes { get; set; }
    /// <summary>Maximum manifest bytes.</summary>
    [Parameter]
    [ValidateRange(1, long.MaxValue)]
    public long MaximumManifestBytes { get; set; }
    /// <summary>Maximum provenance carriers.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int MaximumCarriers { get; set; }
    /// <summary>Maximum container entries.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int MaximumContainerEntries { get; set; }
    /// <summary>Maximum expanded container bytes.</summary>
    [Parameter]
    [ValidateRange(1, long.MaxValue)]
    public long MaximumExpandedContainerBytes { get; set; }
    /// <summary>Maximum embedded assets.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int MaximumEmbeddedAssets { get; set; }
    /// <summary>Maximum text characters inspected.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int MaximumTextCharacters { get; set; }
    /// <summary>Maximum text-integrity findings.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int MaximumFindings { get; set; }
    /// <summary>Skip typographic-space findings.</summary>
    [Parameter]
    public SwitchParameter NoTypographicSpaces { get; set; }
    /// <summary>Skip variation-selector findings.</summary>
    [Parameter]
    public SwitchParameter NoVariationSelectors { get; set; }
    /// <summary>Maximum verifier runtime in seconds; defaults to 30.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int VerificationTimeoutSeconds { get; set; }
    /// <summary>Allow the explicitly configured verifier to resolve remote manifests and trust material.</summary>
    [Parameter]
    public SwitchParameter AllowNetworkAccess { get; set; }
    /// <summary>Include the bounded verifier JSON report.</summary>
    [Parameter]
    public SwitchParameter IncludeRawReport { get; set; }
    /// <summary>Local PEM trust-anchor list.</summary>
    [Parameter]
    public string? TrustAnchorsPath { get; set; }
    /// <summary>Local PEM allowed-certificate list.</summary>
    [Parameter]
    public string? AllowedListPath { get; set; }
    /// <summary>Local C2PA trust configuration.</summary>
    [Parameter]
    public string? TrustConfigurationPath { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var policy = new OfficeProvenanceAssessmentOptions();
        if (MyInvocation.BoundParameters.ContainsKey(nameof(NoTextIntegrity))) policy.InspectTextIntegrity = !NoTextIntegrity.IsPresent;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(NoEmbeddedAssets))) policy.Structural.ProcessEmbeddedAssets = !NoEmbeddedAssets.IsPresent;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(MaximumAssetBytes))) policy.Structural.MaxAssetBytes = MaximumAssetBytes;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(MaximumManifestBytes))) policy.Structural.MaxManifestBytes = MaximumManifestBytes;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(MaximumCarriers))) policy.Structural.MaxCarriers = MaximumCarriers;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(MaximumContainerEntries))) policy.Structural.MaxContainerEntries = MaximumContainerEntries;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(MaximumExpandedContainerBytes))) policy.Structural.MaxExpandedContainerBytes = MaximumExpandedContainerBytes;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(MaximumEmbeddedAssets))) policy.Structural.MaxEmbeddedAssets = MaximumEmbeddedAssets;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(MaximumTextCharacters))) policy.TextIntegrity.MaxCharacters = MaximumTextCharacters;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(MaximumFindings))) policy.TextIntegrity.MaxFindings = MaximumFindings;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(NoTypographicSpaces))) policy.TextIntegrity.IncludeTypographicSpaces = !NoTypographicSpaces.IsPresent;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(NoVariationSelectors))) policy.TextIntegrity.IncludeVariationSelectors = !NoVariationSelectors.IsPresent;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(VerificationTimeoutSeconds))) policy.Verification.Timeout = TimeSpan.FromSeconds(VerificationTimeoutSeconds);
        if (MyInvocation.BoundParameters.ContainsKey(nameof(AllowNetworkAccess))) policy.Verification.AllowNetworkAccess = AllowNetworkAccess.IsPresent;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(IncludeRawReport))) policy.Verification.IncludeRawReport = IncludeRawReport.IsPresent;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(TrustAnchorsPath))) policy.Verification.TrustAnchorsPath = ResolveTrustPath(TrustAnchorsPath);
        if (MyInvocation.BoundParameters.ContainsKey(nameof(AllowedListPath))) policy.Verification.AllowedListPath = ResolveTrustPath(AllowedListPath);
        if (MyInvocation.BoundParameters.ContainsKey(nameof(TrustConfigurationPath))) policy.Verification.TrustConfigurationPath = ResolveTrustPath(TrustConfigurationPath);
        WriteObject(policy);
    }

    private string? ResolveTrustPath(string? path) {
        if (path == null) return null;
        if (Uri.TryCreate(path, UriKind.Absolute, out Uri? uri)
            && (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps)) return path;
        return OfficeFileCommandUtilities.ResolvePath(this, path);
    }
}
