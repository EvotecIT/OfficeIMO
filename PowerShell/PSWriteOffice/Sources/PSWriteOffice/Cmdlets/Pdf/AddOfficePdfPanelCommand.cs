using PSWriteOffice.Foundation;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Adds a visually separated panel paragraph to a PDF document.</summary>
/// <example>
///   <summary>Add a callout panel to a report.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePdf -Path .\Examples\Documents\PdfPanel.pdf {
///     Add-OfficePdfHeading -Text 'Executive summary'
///     Add-OfficePdfPanel -Text 'No critical incidents were detected in the current reporting window.' -Alignment Center
/// }</code>
///   <para>Adds a highlighted panel paragraph to the generated PDF.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficePdfPanel", DefaultParameterSetName = "Text")]
[OutputType(typeof(PdfDocument))]
[Alias("PdfPanel") ]
public sealed class AddOfficePdfPanelCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>PDF document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public PdfDocument Document { get; set; } = null!;

    /// <summary>Panel text.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Text")]
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Document")]
    public string Text { get; set; } = string.Empty;

    /// <summary>Rich PdfText child content.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Content")]
    public ScriptBlock? Content { get; set; }

    /// <summary>Panel alignment.</summary>
    [Parameter]
    public OfficeAlignment Alignment { get; set; } = OfficeAlignment.Left;

    /// <summary>Emit the updated document.</summary>
    [Parameter(ParameterSetName = ParameterSetDocument)]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var runs = new List<PdfTextRun>();
        if (Content != null) { using (OfficeDocumentContext.Push(runs)) Content.InvokeReturnAsIs(); }
        else runs.Add(PdfTextRun.Normal(Text));
        var document = PdfCommandUtilities.ComposeContent(this, Document, ParameterSetName, ParameterSetDocument,
            content => content.AddPanelParagraph(p => p.AddRuns(runs), align: Alignment));
        if (PassThru.IsPresent && document != null)
        {
            WriteObject(document);
        }
    }
}
