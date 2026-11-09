using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Copy PDF pages through the native PDF engine.</summary>
[Cmdlet(VerbsCommon.Copy, "OfficePdfPage", DefaultParameterSetName = "Document", SupportsShouldProcess = true)]
[OutputType(typeof(PdfDocument), typeof(FileInfo))]
public sealed class CopyOfficePdfPageCommand : PdfMutationCmdlet {
    /// <summary>Caller-ordered one-based pages or range expression.</summary>
    [Parameter(Mandatory = true), OfficePageRangeTransformation, ValidateNotNull]
    public OfficePageRange Page { get; set; } = null!;
    /// <inheritdoc />
    protected override string Operation => "Copy PDF pages";
    /// <inheritdoc />
    protected override PdfDocument Apply(PdfDocument document, CancellationToken cancellationToken) =>
        document.Pages.Extract(Page.Selection, cancellationToken);
}
