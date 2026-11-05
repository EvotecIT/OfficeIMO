using PSWriteOffice.Foundation;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Adds invisible vertical spacing to a generated PDF document.</summary>
/// <example>
///   <summary>Add vertical rhythm between sections.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePdf -Path .\Examples\Documents\PdfSpacer.pdf {
///     Add-OfficePdfHeading -Text 'Summary'
///     Add-OfficePdfParagraph -Text 'First block.'
///     Add-OfficePdfSpacer -Height 18
///     Add-OfficePdfParagraph -Text 'Second block after additional spacing.'
/// }</code>
///   <para>Adds whitespace without adding visible content.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficePdfSpacer", DefaultParameterSetName = ParameterSetContext)]
[OutputType(typeof(PdfDocument))]
[Alias("PdfSpacer") ]
public sealed class AddOfficePdfSpacerCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>PDF document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public PdfDocument Document { get; set; } = null!;

    /// <summary>Vertical space height in PDF points.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [OfficeLengthTransformation]
    public OfficeLength Height { get; set; }

    /// <summary>Emit the updated document.</summary>
    [Parameter(ParameterSetName = ParameterSetDocument)]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (Height.ToEmus() < 0) throw new PSArgumentOutOfRangeException(nameof(Height), Height, "Spacer height cannot be negative.");
        var document = PdfCommandUtilities.ComposeContent(this, Document, ParameterSetName, ParameterSetDocument,
            content => content.Spacer(Height));
        if (PassThru.IsPresent && document != null)
        {
            WriteObject(document);
        }
    }
}
