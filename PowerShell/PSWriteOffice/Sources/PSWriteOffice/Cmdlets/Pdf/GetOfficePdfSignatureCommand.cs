using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Gets lightweight PDF signature structure and preservation validation.</summary>
/// <remarks>
/// Reports signature fields, byte-range structure, DocMDP permissions, DSS/LTV evidence, and append-only preservation markers.
/// Certificate-chain trust, revocation, digest, and CMS cryptographic verification are intentionally not performed.
/// </remarks>
/// <example>
///   <summary>Inspect signatures before preserving or migrating a PDF.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$report = Get-OfficePdfSignature -Path .\Signed.pdf
/// $report.Signatures
/// $report.Findings</code>
///   <para>Reads signature structure and reports whether OfficeIMO.Pdf found structural issues.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficePdfSignature")]
[OutputType(typeof(PdfSignatureValidationReport))]
public sealed class GetOfficePdfSignatureCommand : PSCmdlet
{
    /// <summary>PDF file path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Password used to inspect a Standard password-encrypted PDF.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <summary>After successful password authentication, explicitly ignore owner-imposed usage restrictions.</summary>
    [Parameter]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        WriteObject(PdfDocument
            .Load(
                PdfCommandUtilities.ResolvePath(this, Path),
                PdfCommandUtilities.CreateReadOptions(Password, IgnorePermissionRestrictions.IsPresent))
            .Security.ValidateSignatures());
    }
}
