using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Adds a plain or rich PDF paragraph using PdfText children.</summary>
/// <example><summary>Mix plain and emphasized text.</summary>
/// <code>New-OfficePdf ./Report.pdf { PdfParagraph { PdfText 'Status: '; PdfText 'Healthy' -Bold } }</code></example>
[Cmdlet(VerbsCommon.Add, "OfficePdfParagraph", DefaultParameterSetName = "Text")]
[OutputType(typeof(PdfDocument))]
[Alias("PdfParagraph") ]
public sealed class AddOfficePdfParagraphCommand : PSCmdlet {
    /// <summary>Live PDF outside a composition block.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "DocumentText")]
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "DocumentContent")]
    public PdfDocument? Document { get; set; }
    /// <summary>Plain paragraph text, including an empty paragraph.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Text")]
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "DocumentText"), AllowEmptyString]
    public string Text { get; set; } = string.Empty;
    /// <summary>Rich paragraph composed with PdfText.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Content")]
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "DocumentContent")]
    public ScriptBlock? Content { get; set; }
    /// <summary>Paragraph alignment.</summary>
    [Parameter] public OfficeAlignment Alignment { get; set; } = OfficeAlignment.Left;
    /// <summary>Default paragraph color.</summary>
    [Parameter, OfficeColorTransformation, ArgumentCompleter(typeof(OfficeColorCompleter))] public OfficeColor? Color { get; set; }
    /// <summary>Emit the updated live document.</summary>
    [Parameter(ParameterSetName = "DocumentText"), Parameter(ParameterSetName = "DocumentContent")]
    public SwitchParameter PassThru { get; set; }
    /// <inheritdoc />
    protected override void ProcessRecord() {
        var runs = new System.Collections.Generic.List<PdfTextRun>();
        if (Content != null) {
            using (OfficeDocumentContext.Push(runs)) Content.InvokeReturnAsIs();
        } else runs.Add(new PdfTextRun(Text));
        var document = PdfCommandUtilities.ComposeContent(this, Document, Document == null ? "Context" : "Document", "Document",
            content => content.AddParagraph(paragraph => paragraph.AddRuns(runs), Alignment, Color));
        if (PassThru && document != null) WriteObject(document);
    }
}
