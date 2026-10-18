using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Adds a horizontal rule divider to a generated PDF document.</summary>
/// <example>
///   <summary>Separate report sections with a divider.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePdf -Path .\Examples\Documents\PdfDivider.pdf {
///     Add-OfficePdfHeading -Text 'Executive summary'
///     Add-OfficePdfParagraph -Text 'The service is healthy.'
///     Add-OfficePdfHorizontalRule -Color '#CBD5E1' -Thickness 0.75 -SpacingBefore 10 -SpacingAfter 10
///     Add-OfficePdfHeading -Text 'Signals' -Level 2
///   }</code>
///   <para>Adds a visual divider between generated PDF sections.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficePdfHorizontalRule", DefaultParameterSetName = ParameterSetContext)]
[Alias("PdfHorizontalRule", "PdfHr")]
[OutputType(typeof(PdfDocument))]
public sealed class AddOfficePdfHorizontalRuleCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>PDF document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public PdfDocument Document { get; set; } = null!;

    /// <summary>Rule thickness in PDF points.</summary>
    [Parameter]
    public double? Thickness { get; set; }

    /// <summary>Rule color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? Color { get; set; }

    /// <summary>Spacing before the rule in PDF points.</summary>
    [Parameter]
    public double? SpacingBefore { get; set; }

    /// <summary>Spacing after the rule in PDF points.</summary>
    [Parameter]
    public double? SpacingAfter { get; set; }

    /// <summary>Emit the updated document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var document = PdfCommandUtilities.ComposeContent(this, Document, ParameterSetName, ParameterSetDocument,
            content => content.HR(Thickness, PdfCommandUtilities.ParseColor(Color), SpacingBefore, SpacingAfter));
        if (PassThru.IsPresent && document != null)
        {
            WriteObject(document);
        }
    }
}
