using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Adds a bullet or numbered list to a PDF document.</summary>
/// <example>
///   <summary>Add action items to a report.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePdf -Path .\Examples\Documents\PdfList.pdf {
///     Add-OfficePdfHeading -Text 'Next actions'
///     Add-OfficePdfList -Items 'Confirm owner','Publish summary','Schedule review' -Numbered
/// }</code>
///   <para>Adds a numbered list in the generated PDF flow.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficePdfList", DefaultParameterSetName = ParameterSetContext)]
[Alias("PdfList")]
[OutputType(typeof(PdfDocument))]
public sealed class AddOfficePdfListCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>PDF document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public PdfDocument Document { get; set; } = null!;

    /// <summary>List item text.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string[] Items { get; set; } = System.Array.Empty<string>();

    /// <summary>Create a numbered list instead of a bullet list.</summary>
    [Parameter]
    public SwitchParameter Numbered { get; set; }

    /// <summary>Number to use for the first numbered item.</summary>
    [Parameter]
    public int StartNumber { get; set; } = 1;

    /// <summary>List alignment.</summary>
    [Parameter]
    public PdfAlign Align { get; set; } = PdfAlign.Left;

    /// <summary>Emit the updated document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var document = PdfCommandUtilities.ComposeContent(this, Document, ParameterSetName, ParameterSetDocument, content =>
        {
            if (Numbered.IsPresent)
            {
                content.Numbered(Items, Align, startNumber: StartNumber);
            }
            else
            {
                content.Bullets(Items, Align);
            }
        });

        if (PassThru.IsPresent && document != null)
        {
            WriteObject(document);
        }
    }
}
