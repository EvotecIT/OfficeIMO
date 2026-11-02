using System.Management.Automation;
using OfficeIMO.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Completes a live PDF document's lifetime without saving.</summary>
/// <example><summary>Finish using a PDF without writing pending changes.</summary>
/// <code>$pdf | Close-OfficePdf</code></example>
[Cmdlet(VerbsCommon.Close, "OfficePdf", DefaultParameterSetName = "Document")]
public sealed class CloseOfficePdfCommand : PSCmdlet {
    /// <summary>Live PDF document to close.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true), ValidateNotNull]
    public PdfDocument Document { get; set; } = null!;

    /// <inheritdoc />
    protected override void ProcessRecord() => Document.Dispose();
}
