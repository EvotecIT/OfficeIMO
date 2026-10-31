using System;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Chooses native presentation inspection reports without changing a document.</summary>
/// <example><summary>Add feature and review inventories to a detailed inspection.</summary>
/// <code>$options = New-OfficePowerPointInspectionOptions -Features -ReviewComments
/// Get-OfficeDocumentInfo -LiteralPath ./Review.pptx -Detailed -Options $options</code></example>
/// <example><summary>Inspect accessibility on an unsaved document.</summary>
/// <code>$options = New-OfficePowerPointInspectionOptions -NoPackageValidation -NoPreflight
/// (Get-OfficeDocumentInfo -Document $deck -Detailed -Options $options).Details.Accessibility</code></example>
[Cmdlet(VerbsCommon.New, "OfficePowerPointInspectionOptions")]
[OutputType(typeof(PowerPointInspectionOptions))]
public sealed class NewOfficePowerPointInspectionOptionsCommand : OfficeCmdlet {
    /// <summary>Gets or sets whether to omit Open XML package validation.</summary>
    [Parameter]
    public SwitchParameter NoPackageValidation { get; set; }
    /// <summary>Gets or sets whether to omit layout, text-fit and asset preflight.</summary>
    [Parameter]
    public SwitchParameter NoPreflight { get; set; }
    /// <summary>Gets or sets whether to omit accessibility inspection.</summary>
    [Parameter]
    public SwitchParameter NoAccessibility { get; set; }
    /// <summary>Gets or sets whether to inventory supported and preserved features.</summary>
    [Parameter]
    public SwitchParameter Features { get; set; }
    /// <summary>Gets or sets whether to inspect review comments.</summary>
    [Parameter]
    public SwitchParameter ReviewComments { get; set; }
    /// <summary>Gets or sets whether to inspect animation metadata.</summary>
    [Parameter]
    public SwitchParameter Animations { get; set; }
    /// <summary>Gets or sets whether to inspect package signature metadata.</summary>
    [Parameter]
    public SwitchParameter Signatures { get; set; }
    /// <summary>Gets or sets whether to create PNG, SVG and snapshot proof for each slide.</summary>
    [Parameter]
    public SwitchParameter Visuals { get; set; }
    /// <summary>Gets or sets the Open XML version used by package validation; defaults to Microsoft 365.</summary>
    [Parameter]
    public OfficeOpenXmlFileFormatVersion FileFormatVersion { get; set; } = OfficeOpenXmlFileFormatVersion.Microsoft365;

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (!Enum.IsDefined(typeof(OfficeOpenXmlFileFormatVersion), FileFormatVersion)) {
            InvalidArgument(new ArgumentException("Choose a supported Office file format version."), FileFormatVersion);
            return;
        }
        WriteObject(new PowerPointInspectionOptions {
            ValidatePackage = FeatureEnabled("PackageValidation"),
            InspectPreflight = FeatureEnabled("Preflight"),
            InspectAccessibility = FeatureEnabled("Accessibility"),
            InspectFeatures = Features.IsPresent,
            InspectReviewComments = ReviewComments.IsPresent,
            InspectAnimations = Animations.IsPresent,
            InspectSignatures = Signatures.IsPresent,
            InspectVisuals = Visuals.IsPresent,
            FileFormatVersion = FileFormatVersion
        });
    }
}
