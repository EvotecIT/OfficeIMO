using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Clears generated PDF page background shapes.</summary>
/// <example>
///   <summary>Remove queued background shapes during composition.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePdf -Path .\Examples\Documents\PdfNoBackgroundShape.pdf {
///     Add-OfficePdfBackgroundShape -Shape Rectangle -FillColor '#EEF2FF' -X 0 -Y 0 -Width 595 -Height 120
///     Clear-OfficePdfBackgroundShape
///     Add-OfficePdfHeading -Text 'Clean variant'
/// }</code>
///   <para>Clears generated page background shapes while the PDF page is being composed.</para>
/// </example>
[Cmdlet(VerbsCommon.Clear, "OfficePdfBackgroundShape", DefaultParameterSetName = ParameterSetContext)]
[OutputType(typeof(PdfDocument))]
public sealed class ClearOfficePdfBackgroundShapeCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>Compatibility parameter. Page composition is supported only inside New-OfficePdf.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public PdfDocument Document { get; set; } = null!;

    /// <summary>Emit the updated document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var document = PdfCommandUtilities.ComposePage(this, Document, ParameterSetName, ParameterSetDocument,
            page => page.ClearBackgroundShapes());
        if (PassThru.IsPresent && document != null)
        {
            WriteObject(document);
        }
    }
}
