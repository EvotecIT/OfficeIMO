using System.Management.Automation;
using OfficeIMO.OpenDocument;
using OfficeIMO.PowerPoint.OpenDocument;

namespace PSWriteOffice.Cmdlets.OpenDocument;

/// <summary>Creates PowerPoint/OpenDocument conversion settings.</summary>
/// <example>
///   <summary>Include slide images, notes, and basic formatting.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$options = New-OfficePowerPointOpenDocumentOptions -IncludeImages -IncludeSpeakerNotes -IncludeBasicFormatting
/// ConvertTo-OfficeOpenDocument -Path .\Deck.pptx -OutputPath .\Deck.odp -PowerPointOptions $options</code>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficePowerPointOpenDocumentOptions")]
[OutputType(typeof(PowerPointOpenDocumentConversionOptions))]
public sealed class NewOfficePowerPointOpenDocumentOptionsCommand : PSCmdlet {
    /// <summary>Whether conversion loss is reported or rejected.</summary>
    [Parameter] public OdfConversionLossPolicy? LossPolicy { get; set; }
    /// <summary>Copy supported embedded images.</summary>
    [Parameter] public SwitchParameter IncludeImages { get; set; }
    /// <summary>Copy plain speaker-note text.</summary>
    [Parameter] public SwitchParameter IncludeSpeakerNotes { get; set; }
    /// <summary>Copy common fills, outlines, and text-run formatting.</summary>
    [Parameter] public SwitchParameter IncludeBasicFormatting { get; set; }
    /// <summary>Maximum rows in converted presentation tables.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int? MaxTableRows { get; set; }
    /// <summary>Maximum columns in converted presentation tables.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int? MaxTableColumns { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new PowerPointOpenDocumentConversionOptions();
        if (LossPolicy.HasValue) options.LossPolicy = LossPolicy.Value;
        Apply(nameof(IncludeImages), value => options.IncludeImages = value);
        Apply(nameof(IncludeSpeakerNotes), value => options.IncludeSpeakerNotes = value);
        Apply(nameof(IncludeBasicFormatting), value => options.IncludeBasicFormatting = value);
        if (MaxTableRows.HasValue) options.MaxTableRows = MaxTableRows.Value;
        if (MaxTableColumns.HasValue) options.MaxTableColumns = MaxTableColumns.Value;
        WriteObject(options);
    }
    private void Apply(string name, System.Action<bool> setter) {
        if (!MyInvocation.BoundParameters.ContainsKey(name)) return;
        setter(((SwitchParameter)MyInvocation.BoundParameters[name]).IsPresent);
    }
}
