using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Adds a bulleted or numbered list built from PdfListItem children.</summary>
/// <example><summary>Use the native OfficePdfList command.</summary><code>New-OfficePdf -Path .\Actions.pdf {
///     PdfList {
///         PdfListItem 'Confirm the service owner'
///         PdfListItem 'Publish the summary'
///         PdfListItem 'Schedule the review'
///     }
/// }</code></example>
[Cmdlet("Add", "OfficePdfList", DefaultParameterSetName = "Context")]
[OutputType(typeof(PdfDocument))]
[Alias("PdfList") ]
public sealed class AddOfficePdfListCommand : OfficeCmdlet {
    /// <summary>Content for this operation.</summary>
    [Parameter(ParameterSetName = "Context", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "Document", Mandatory = true, Position = 0)]
    public ScriptBlock? Content { get; set; }
    /// <summary>Numbered for this operation.</summary>
    [Parameter(ParameterSetName = "Context")]
    [Parameter(ParameterSetName = "Items")]
    [Parameter(ParameterSetName = "Document")]
    [Parameter(ParameterSetName = "DocumentItems")]
    public SwitchParameter Numbered { get; set; }
    /// <summary>StartNumber for this operation.</summary>
    [Parameter(ParameterSetName = "Context")]
    [Parameter(ParameterSetName = "Items")]
    [Parameter(ParameterSetName = "Document")]
    [Parameter(ParameterSetName = "DocumentItems")]
    public int? StartNumber { get; set; }
    /// <summary>Alignment for this operation.</summary>
    [Parameter(ParameterSetName = "Context")]
    [Parameter(ParameterSetName = "Items")]
    [Parameter(ParameterSetName = "Document")]
    [Parameter(ParameterSetName = "DocumentItems")]
    public OfficeAlignment? Alignment { get; set; }
    /// <summary>Items for this operation.</summary>
    [Parameter(ParameterSetName = "Items", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "DocumentItems", Mandatory = true, Position = 0)]
    public string[]? Items { get; set; }
    /// <summary>Document for this operation.</summary>
    [Parameter(ParameterSetName = "Document", Mandatory = true, ValueFromPipeline = true)]
    [Parameter(ParameterSetName = "DocumentItems", Mandatory = true, ValueFromPipeline = true)]
    public PdfDocument? Document { get; set; }
    /// <summary>Document set only; inside a block the document is returned by New-OfficePdf.</summary>
    [Parameter(ParameterSetName = "Document")]
    [Parameter(ParameterSetName = "DocumentItems")]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var items = new List<PdfListItem>();
        if (Content != null) { using (OfficeDocumentContext.Push(items)) Content.InvokeReturnAsIs(); }
        else foreach (string item in Items!) items.Add(new PdfListItem(item));
        var document = PdfCommandUtilities.ComposeContent(this, Document, Document == null ? "Context" : "Document", "Document",
            content => { if (Numbered) content.AddNumberedList(items, Alignment ?? OfficeAlignment.Left, startNumber: StartNumber ?? 1); else content.AddList(items, Alignment ?? OfficeAlignment.Left); });
        if (PassThru && document != null) WriteObject(document);
    }
}
