using System.Collections;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Creates PDF encryption settings; Protect-OfficeDocument supplies the actual passwords.</summary>
/// <example><summary>Select AES-256 while leaving metadata readable.</summary><code>$options = New-OfficePdfEncryptionOptions -Algorithm Aes256 -NoMetadataEncryption
/// Protect-OfficeDocument -Path ./Report.pdf -Encrypt -Password $password -Options $options</code></example>
[Cmdlet(VerbsCommon.New, "OfficePdfEncryptionOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(Hashtable))]
public sealed class NewOfficePdfEncryptionOptionsCommand : OfficeCmdlet {
    /// <summary>Native PDF Standard encryption algorithm.</summary>
    [Parameter(ParameterSetName = "Default")]
    public PdfStandardEncryptionAlgorithm Algorithm { get; set; } = PdfStandardEncryptionAlgorithm.Aes256;
    /// <summary>Leave document metadata unencrypted.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter NoMetadataEncryption { get; set; }
    /// <inheritdoc />
    protected override void ProcessRecord() {
        WriteObject(new Hashtable(System.StringComparer.OrdinalIgnoreCase) {
            ["Algorithm"] = Algorithm,
            ["EncryptMetadata"] = !NoMetadataEncryption.IsPresent
        });
    }
}
