using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Reports whether OfficeIMO.Pdf can read or rewrite a PDF safely.</summary>
/// <example>
///   <summary>Preflight a PDF before migration operations.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$preflight = Get-OfficePdfPreflight -Path .\Examples\Documents\Report.pdf
/// $preflight.HasReadBlockers
/// $preflight.HasRewriteBlockers</code>
///   <para>Checks whether OfficeIMO.Pdf can read or rewrite the PDF safely.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficePdfPreflight")]
[OutputType(typeof(PdfDocumentPreflight))]
public sealed class GetOfficePdfPreflightCommand : PSCmdlet
{
    /// <summary>PDF file path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Password used to preflight a Standard password-encrypted PDF.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <summary>After successful password authentication, explicitly ignore owner-imposed usage restrictions.</summary>
    [Parameter]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        WriteObject(PdfDocument.Preflight(
            PdfCommandUtilities.ResolvePath(this, Path),
            PdfCommandUtilities.CreateReadOptions(Password, IgnorePermissionRestrictions.IsPresent)));
    }
}
