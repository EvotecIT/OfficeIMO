using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Adds a page break to a PDF document.</summary>
/// <example>
///   <summary>Start an appendix on a new page.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePdf -Path .\Examples\Documents\PdfPageBreak.pdf {
///     Add-OfficePdfHeading -Text 'Service review'
///     Add-OfficePdfParagraph -Text 'Summary content stays on the first page.'
///     Add-OfficePdfPageBreak
///     Add-OfficePdfHeading -Text 'Appendix' -Level 2
/// }</code>
///   <para>Forces the appendix section to begin on the next page.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficePdfPageBreak", DefaultParameterSetName = ParameterSetContext)]
[Alias("PdfPageBreak")]
[OutputType(typeof(PdfDocument))]
public sealed class AddOfficePdfPageBreakCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>PDF document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public PdfDocument Document { get; set; } = null!;

    /// <summary>Emit the updated document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var document = PdfCommandUtilities.ComposeContent(this, Document, ParameterSetName, ParameterSetDocument,
            content => content.PageBreak());
        if (PassThru.IsPresent && document != null)
        {
            WriteObject(document);
        }
    }
}
