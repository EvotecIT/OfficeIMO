using System.IO;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Injects externally produced CMS, CAdES, or timestamp signature bytes into a prepared PDF signature placeholder.</summary>
/// <example>
///   <summary>Inject a detached CMS signature into a prepared PDF.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Set-OfficePdfSignature -Path .\Prepared.pdf -SignaturePath .\signature.der -OutputPath .\Signed.pdf
/// Get-OfficePdfSignature -Path .\Signed.pdf</code>
///   <para>Writes a PDF with the reserved /Contents hex slot patched in place.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficePdfSignature", SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo), typeof(PdfSignatureValidationReport))]
public sealed class SetOfficePdfSignatureCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    /// <summary>Prepared PDF path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>DER/CMS/CAdES/TSA response bytes to inject into the reserved /Contents slot.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string SignaturePath { get; set; } = string.Empty;

    /// <summary>Output signed PDF path.</summary>
    [Parameter(Mandatory = true, Position = 2)]
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>Password used to authenticate an encrypted prepared PDF.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <summary>After successful password authentication, explicitly ignore owner-imposed signature restrictions.</summary>
    [Parameter]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }

    /// <summary>Return a signature validation report for the written PDF instead of only the output file.</summary>
    [Parameter]
    public SwitchParameter PassThruReport { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var inputPath = PdfCommandUtilities.ResolvePath(this, Path);
        var signaturePath = PdfCommandUtilities.ResolvePath(this, SignaturePath);
        var outputPath = PdfCommandUtilities.ResolvePath(this, OutputPath);
        if (!PdfCommandUtilities.ShouldWrite(this, outputPath, "Write signed PDF")) {
            return;
        }

        PdfCommandUtilities.EnsureDirectory(outputPath);

        var document = PdfDocument
            .Load(inputPath, PdfCommandUtilities.CreateReadOptions(Password, IgnorePermissionRestrictions.IsPresent))
            .Security.CompleteExternalSignature(File.ReadAllBytes(signaturePath));
        document.Save(outputPath).RequireSuccess();
        if (PassThruReport.IsPresent) {
            WriteObject(document.Security.ValidateSignatures());
            return;
        }

        WritePassThru(new FileInfo(outputPath));
    }
}
