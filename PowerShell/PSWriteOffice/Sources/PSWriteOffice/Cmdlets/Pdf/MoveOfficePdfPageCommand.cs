using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Move PDF pages through the native PDF engine.</summary>
[Cmdlet(VerbsCommon.Move, "OfficePdfPage", DefaultParameterSetName = "Document", SupportsShouldProcess = true)]
[OutputType(typeof(PdfDocument), typeof(FileInfo))]
public sealed class MoveOfficePdfPageCommand : PdfMutationCmdlet {
    /// <summary>Caller-ordered one-based pages or range expression.</summary>
    [Parameter(Mandatory = true), OfficePageRangeTransformation, ValidateNotNull]
    public OfficePageRange Page { get; set; } = null!;
    /// <summary>One-based page before which selected pages are inserted; page count plus one appends.</summary>
    [Parameter(Mandatory = true), ValidateRange(1, int.MaxValue)]
    public int BeforePage { get; set; }
    /// <inheritdoc />
    protected override string Operation => "Move PDF pages";
    /// <inheritdoc />
    protected override PdfDocument Apply(PdfDocument document, CancellationToken cancellationToken) =>
        document.Pages.Move(BeforePage, Page.Selection, cancellationToken);
}
