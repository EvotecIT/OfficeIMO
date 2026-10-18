using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Adds a paragraph to a PDF document.</summary>
/// <example>
///   <summary>Add body text to a generated PDF.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePdf -Path .\Examples\Documents\PdfParagraph.pdf {
///     Add-OfficePdfHeading -Text 'Status'
///     Add-OfficePdfParagraph -Text 'All monitored services are currently healthy.' -Color '#166534'
/// }</code>
///   <para>Adds a colored body paragraph after a heading.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficePdfParagraph", DefaultParameterSetName = ParameterSetContext)]
[Alias("PdfParagraph")]
[OutputType(typeof(PdfDocument))]
public sealed class AddOfficePdfParagraphCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>PDF document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public PdfDocument Document { get; set; } = null!;

    /// <summary>Paragraph text.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Text { get; set; } = string.Empty;

    /// <summary>Paragraph alignment.</summary>
    [Parameter]
    public PdfAlign Align { get; set; } = PdfAlign.Left;

    /// <summary>Optional text color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? Color { get; set; }

    /// <summary>Emit the updated document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var document = PdfCommandUtilities.ComposeContent(this, Document, ParameterSetName, ParameterSetDocument,
            content => content.Paragraph(p => p.Text(Text), Align, PdfCommandUtilities.ParseColor(Color)));
        if (PassThru.IsPresent && document != null)
        {
            WriteObject(document);
        }
    }
}
