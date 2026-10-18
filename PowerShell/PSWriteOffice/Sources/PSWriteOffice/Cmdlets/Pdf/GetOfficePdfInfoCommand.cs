using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Gets PDF metadata, page information, forms, links, and structural flags.</summary>
/// <remarks>
/// The returned OfficeIMO.Pdf inspection object is useful for validating generated artifacts, migration scripts,
/// and existing PDFs before follow-up operations such as splitting, stamping, or metadata updates.
/// </remarks>
/// <example>
///   <summary>Inspect a generated PDF.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$info = Get-OfficePdfInfo -Path .\Report.pdf
/// $info.PageCount
/// $info.LinkUris</code>
///   <para>Reads page count and link information from the PDF.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficePdfInfo")]
[OutputType(typeof(PdfDocumentInfo))]
public sealed class GetOfficePdfInfoCommand : PSCmdlet
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
            .Inspect());
    }
}
