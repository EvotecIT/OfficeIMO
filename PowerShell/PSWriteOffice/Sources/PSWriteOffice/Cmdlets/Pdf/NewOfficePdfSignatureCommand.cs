using System.IO;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Prepares an existing PDF for external digital signing by appending a signature field, /ByteRange, and reserved /Contents placeholder.</summary>
/// <remarks>
/// The command does not create CMS, CAdES, timestamp, certificate-chain, or revocation data. Use the returned byte range or digest with an external signing service, then inject the produced signature bytes with Set-OfficePdfSignature.
/// </remarks>
/// <example>
///   <summary>Prepare a PDF for detached CMS signing.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$plan = New-OfficePdfSignature -Path .\Input.pdf -OutputPath .\Prepared.pdf -FieldName Approval -Name 'Alice' -Reason Approval -PassThruReport
/// $plan.ByteRangeValues
/// $plan.ComputeSha256Digest()</code>
///   <para>Writes a prepared PDF and returns the OfficeIMO.Pdf external signing preparation report.</para>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficePdfSignature", SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo), typeof(PdfExternalSignaturePreparation))]
public sealed class NewOfficePdfSignatureCommand : PSCmdlet
{
    /// <summary>Input PDF path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Output prepared PDF path.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>Password used to authenticate an encrypted PDF.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <summary>After successful password authentication, explicitly ignore owner-imposed signature-field restrictions.</summary>
    [Parameter]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }

    /// <summary>Signature field name to append.</summary>
    [Parameter]
    public string FieldName { get; set; } = "Signature1";

    /// <summary>Signature handler filter name. The default is Adobe.PPKLite.</summary>
    [Parameter]
    public string Filter { get; set; } = "Adobe.PPKLite";

    /// <summary>Signature subfilter that describes the external signature bytes to inject later.</summary>
    [Parameter]
    public PdfExternalSignatureSubFilter SubFilter { get; set; } = PdfExternalSignatureSubFilter.DetachedCms;

    /// <summary>Display signer name stored in the signature dictionary.</summary>
    [Parameter]
    public string? Name { get; set; }

    /// <summary>Signing reason stored in the signature dictionary.</summary>
    [Parameter]
    public string? Reason { get; set; }

    /// <summary>Signing location stored in the signature dictionary.</summary>
    [Parameter]
    public string? Location { get; set; }

    /// <summary>Signer contact information stored in the signature dictionary.</summary>
    [Parameter]
    public string? ContactInfo { get; set; }

    /// <summary>Raw signature bytes to reserve in /Contents before hex encoding.</summary>
    [Parameter]
    public int ReservedBytes { get; set; } = 32768;

    /// <summary>Return the OfficeIMO.Pdf preparation report instead of only the output file.</summary>
    [Parameter]
    public SwitchParameter PassThruReport { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var inputPath = PdfCommandUtilities.ResolvePath(this, Path);
        var outputPath = PdfCommandUtilities.ResolvePath(this, OutputPath);
        if (!PdfCommandUtilities.ShouldWrite(this, outputPath, "Write externally signable PDF"))
        {
            return;
        }

        PdfCommandUtilities.EnsureDirectory(outputPath);

        var options = new PdfExternalSignatureOptions
        {
            FieldName = FieldName,
            Filter = Filter,
            SubFilter = SubFilter,
            Name = Name,
            Reason = Reason,
            Location = Location,
            ContactInfo = ContactInfo,
            ReservedSignatureContentsBytes = ReservedBytes
        };

        var readOptions = PdfCommandUtilities.CreateReadOptions(Password, IgnorePermissionRestrictions.IsPresent);
        PdfExternalSignaturePreparation preparation = PdfDocument
            .Load(inputPath, readOptions)
            .Security.PrepareExternalSignature(options);
        PdfDocument.Load(preparation.PreparedPdf, readOptions).Save(outputPath).RequireSuccess();
        WriteObject(PassThruReport.IsPresent ? preparation : new FileInfo(outputPath));
    }
}
