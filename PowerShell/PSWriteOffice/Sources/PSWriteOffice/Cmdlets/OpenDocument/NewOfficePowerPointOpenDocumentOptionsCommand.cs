using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.PowerPoint.OpenDocument;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.OpenDocument;

/// <summary>Creates PowerPoint/OpenDocument conversion settings.</summary>
/// <example><summary>Omit notes from an OpenDocument presentation.</summary>
/// <code>$options = New-OfficePowerPointOpenDocumentOptions -NoSpeakerNotes
/// Convert-OfficeDocument -Path ./Review.pptx -To OpenDocument -Options $options</code></example>
[Cmdlet(VerbsCommon.New, "OfficePowerPointOpenDocumentOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(PowerPointOpenDocumentConversionOptions))]
public sealed class NewOfficePowerPointOpenDocumentOptionsCommand : OfficeCmdlet {
    /// <summary>Gets or sets whether known omitted content is rejected or reported; omission retains the adapter default.</summary>
    [Parameter(ParameterSetName = "Default")]
    public OfficeConversionLossPolicy? LossPolicy { get; set; }
    /// <summary>Gets or sets whether embedded images are omitted.</summary>
    [Parameter]
    public SwitchParameter NoImages { get; set; }
    /// <summary>Gets or sets whether speaker notes are omitted.</summary>
    [Parameter]
    public SwitchParameter NoSpeakerNotes { get; set; }
    /// <summary>Gets or sets whether common fill, outline and run formatting is omitted.</summary>
    [Parameter]
    public SwitchParameter NoBasicFormatting { get; set; }
    /// <summary>Gets or sets the maximum number of rows in a converted table.</summary>
    [Parameter, ValidateRange(1, int.MaxValue)]
    public int? MaxTableRows { get; set; }
    /// <summary>Gets or sets the maximum number of columns in a converted table.</summary>
    [Parameter, ValidateRange(1, int.MaxValue)]
    public int? MaxTableColumns { get; set; }
    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new PowerPointOpenDocumentConversionOptions {
            IncludeImages = FeatureEnabled("Images"),
            IncludeSpeakerNotes = FeatureEnabled("SpeakerNotes"),
            IncludeBasicFormatting = FeatureEnabled("BasicFormatting")
        };
        if (LossPolicy.HasValue) options.LossPolicy = OpenDocumentCommandUtilities.ResolveLossPolicy(LossPolicy.Value);
        if (MaxTableRows.HasValue) options.MaxTableRows = MaxTableRows.Value;
        if (MaxTableColumns.HasValue) options.MaxTableColumns = MaxTableColumns.Value;
        WriteObject(options);
    }
}
