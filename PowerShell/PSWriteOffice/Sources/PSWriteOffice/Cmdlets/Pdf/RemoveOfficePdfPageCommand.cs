using System.IO;
using System.Management.Automation;
using System.Threading;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Removes selected pages from a live PDF or writes edited file copies.</summary>
/// <para>Inside Open-OfficePdf, omission of Document updates the current document. Outside a block, PassThru returns the native edit for assignment or further piping. File outputs are published atomically.</para>
/// <example><summary>Remove draft pages from an output copy.</summary>
/// <code>Remove-OfficePdfPage -Path ./Report.pdf -Page '4-5' -OutputPath ./Report-Clean.pdf -PassThru</code></example>
/// <example><summary>Remove the last page in an editing block.</summary>
/// <code>Open-OfficePdf ./Report.pdf { Remove-OfficePdfPage -Page last }</code></example>
[Cmdlet(VerbsCommon.Remove, "OfficePdfPage", DefaultParameterSetName = "Document", SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo), typeof(PdfDocument))]
public sealed class RemoveOfficePdfPageCommand : PdfMutationCmdlet {
    /// <summary>Native selector, one-based page or range expression, including last, odd and even.</summary>
    [Parameter(Mandatory = true), PdfPageSelectorTransformation, ValidateNotNull]
    public PdfPageSelector Page { get; set; } = null!;
    /// <inheritdoc />
    protected override string Operation => "Remove PDF pages";
    /// <inheritdoc />
    protected override PdfDocument Apply(PdfDocument document, CancellationToken cancellationToken) =>
        document.Pages.Delete(Page, cancellationToken);
}
