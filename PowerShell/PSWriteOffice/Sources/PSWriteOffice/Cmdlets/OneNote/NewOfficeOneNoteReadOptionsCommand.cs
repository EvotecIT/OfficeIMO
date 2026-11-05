using System.Management.Automation;
using OfficeIMO.OneNote;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.OneNote;

/// <summary>Creates native OneNoteLoadOptions settings for document operations.</summary>
/// <example><summary>Create default settings.</summary><code>$options = New-OfficeOneNoteReadOptions</code></example>
[Cmdlet(VerbsCommon.New, "OfficeOneNoteReadOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(OfficeIMO.OneNote.OneNoteLoadOptions))]
public sealed class NewOfficeOneNoteReadOptionsCommand : OfficeCmdlet {
    /// <summary>MaxInputBytes override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public long? MaxInputBytes { get; set; }
    /// <summary>MaxObjects override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public int? MaxObjects { get; set; }
    /// <summary>MaxAssetBytes override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public long? MaxAssetBytes { get; set; }
    /// <summary>MaxTotalAssetBytes override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public long? MaxTotalAssetBytes { get; set; }
    /// <summary>NoStrictHeaderValidation override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public SwitchParameter NoStrictHeaderValidation { get; set; }
    /// <summary>NoChecksumValidation override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public SwitchParameter NoChecksumValidation { get; set; }
    /// <summary>NoUnknownData override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public SwitchParameter NoUnknownData { get; set; }
    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new OfficeIMO.OneNote.OneNoteLoadOptions();
        if (MaxInputBytes != null) options.MaxInputBytes = MaxInputBytes.Value;
        if (MaxObjects != null) options.MaxObjects = MaxObjects.Value;
        if (MaxAssetBytes != null) options.MaxAssetBytes = MaxAssetBytes.Value;
        if (MaxTotalAssetBytes != null) options.MaxTotalAssetBytes = MaxTotalAssetBytes.Value;
        if (BoundSwitch(nameof(NoStrictHeaderValidation)) is bool valueNoStrictHeaderValidation) options.StrictHeaderValidation = !valueNoStrictHeaderValidation;
        if (BoundSwitch(nameof(NoChecksumValidation)) is bool valueNoChecksumValidation) options.ValidateTransactionChecksums = !valueNoChecksumValidation;
        if (BoundSwitch(nameof(NoUnknownData)) is bool valueNoUnknownData) options.PreserveUnknownData = !valueNoUnknownData;
        WriteObject(options);
    }
}
