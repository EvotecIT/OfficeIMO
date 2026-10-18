using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Adds a heading to a PDF document.</summary>
/// <example>
///   <summary>Create heading levels in a PDF report.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePdf -Path .\Examples\Documents\PdfHeadings.pdf {
///     Add-OfficePdfHeading -Text 'Service Review' -Level 1 -Color '#1D4ED8'
///     Add-OfficePdfHeading -Text 'Open risks' -Level 2
///     Add-OfficePdfParagraph -Text 'Heading levels create the report structure.'
///   }</code>
///   <para>Adds report headings before body content.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficePdfHeading", DefaultParameterSetName = ParameterSetContext)]
[Alias("PdfHeading")]
[OutputType(typeof(PdfDocument))]
public sealed class AddOfficePdfHeadingCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>PDF document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public PdfDocument Document { get; set; } = null!;

    /// <summary>Heading text.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Text { get; set; } = string.Empty;

    /// <summary>Heading level, 1 through 3.</summary>
    [Parameter]
    [ValidateRange(1, 3)]
    public int Level { get; set; } = 1;

    /// <summary>Heading alignment.</summary>
    [Parameter]
    public PdfAlign Align { get; set; } = PdfAlign.Left;

    /// <summary>Optional heading color. Named colors and hexadecimal values are accepted.</summary>
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
        var color = PdfCommandUtilities.ParseColor(Color);
        var document = PdfCommandUtilities.ComposeContent(this, Document, ParameterSetName, ParameterSetDocument, content =>
        {
            switch (Level)
            {
                case 1:
                    content.H1(Text, Align, color);
                    break;
                case 2:
                    content.H2(Text, Align, color);
                    break;
                default:
                    content.H3(Text, Align, color);
                    break;
            }
        });

        if (PassThru.IsPresent && document != null)
        {
            WriteObject(document);
        }
    }
}
