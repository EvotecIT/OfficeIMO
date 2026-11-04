using System.Management.Automation;
using OfficeIMO.OpenDocument;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.OpenDocument;

/// <summary>Creates native OdfSaveOptions settings for document operations.</summary>
/// <example><summary>Create default settings.</summary><code>$options = New-OfficeOpenDocumentSaveOptions</code></example>
[Cmdlet(VerbsCommon.New, "OfficeOpenDocumentSaveOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(OfficeIMO.OpenDocument.OdfSaveOptions))]
public sealed class NewOfficeOpenDocumentSaveOptionsCommand : OfficeCmdlet {
    /// <summary>CompatibilityProfile override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public OdfCompatibilityProfile? CompatibilityProfile { get; set; }
    /// <summary>SignatureHandling override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public OdfSignatureHandling? SignatureHandling { get; set; }
    /// <summary>EncryptionHandling override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public OdfEncryptionHandling? EncryptionHandling { get; set; }
    /// <summary>NoDeterministic override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public SwitchParameter NoDeterministic { get; set; }
    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new OfficeIMO.OpenDocument.OdfSaveOptions();
        if (CompatibilityProfile != null) options.CompatibilityProfile = CompatibilityProfile.Value;
        if (SignatureHandling != null) options.SignatureHandling = SignatureHandling.Value;
        if (EncryptionHandling != null) options.EncryptionHandling = EncryptionHandling.Value;
        if (BoundSwitch(nameof(NoDeterministic)) is bool valueNoDeterministic) options.Deterministic = !valueNoDeterministic;
        WriteObject(options);
    }
}
