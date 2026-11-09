using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Adds one list item, from -Text or from PdfText runs.</summary>
/// <example><summary>Use the native OfficePdfListItem command.</summary><code>New-OfficePdf -Path .\Checklist.pdf {
///     PdfList {
///         PdfListItem 'Backups verified'
///         PdfListItem 'Monitoring silenced'
///     }
/// }</code></example>
[Cmdlet("Add", "OfficePdfListItem", DefaultParameterSetName = "Text")]
[OutputType(typeof(void))]
[Alias("PdfListItem") ]
public sealed class AddOfficePdfListItemCommand : OfficeCmdlet {
    /// <summary>Text for this operation.</summary>
    [Parameter(ParameterSetName = "Text", Mandatory = true, Position = 0)]
    public string? Text { get; set; }
    /// <summary>Content for this operation.</summary>
    [Parameter(ParameterSetName = "Content", Mandatory = true, Position = 0)]
    public ScriptBlock? Content { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var items = OfficeDocumentContext.Find<List<PdfListItem>>() ?? throw new PSInvalidOperationException("Use PdfListItem inside PdfList.");
        if (Content == null) items.Add(new PdfListItem(Text!));
        else {
            var runs = new List<PdfTextRun>();
            using (OfficeDocumentContext.Push(runs)) Content.InvokeReturnAsIs();
            items.Add(new PdfListItem(runs));
        }
    }
}
