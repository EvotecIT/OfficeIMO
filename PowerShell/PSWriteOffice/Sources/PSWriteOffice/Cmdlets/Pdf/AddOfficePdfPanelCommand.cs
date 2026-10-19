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
///     Add-OfficePdfPanel -Text 'No critical incidents were detected in the current reporting window.' -Align Center
/// }</code>
///   <para>Adds a highlighted panel paragraph to the generated PDF.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficePdfPanel", DefaultParameterSetName = ParameterSetContext)]
[Alias("PdfPanel")]
[OutputType(typeof(PdfDocument))]
public sealed class AddOfficePdfPanelCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>PDF document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public PdfDocument Document { get; set; } = null!;

    /// <summary>Panel text.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Text { get; set; } = string.Empty;

    /// <summary>Panel alignment.</summary>
    [Parameter]
    public PdfAlign Align { get; set; } = PdfAlign.Left;

    /// <summary>Emit the updated document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var document = PdfCommandUtilities.ComposeContent(this, Document, ParameterSetName, ParameterSetDocument,
            content => content.PanelParagraph(p => p.Text(Text), align: Align));
        if (PassThru.IsPresent && document != null)
        {
            WriteObject(document);
        }
    }
}
