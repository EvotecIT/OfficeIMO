using System.Management.Automation;
using OfficeIMO.ChartForgeX;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;
using PSWriteOffice.Services.Visuals;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Adds a ChartForgeX artifact, portable SVG, or converted Office visual to PDF flow content.</summary>
[Cmdlet(VerbsCommon.Add, "OfficePdfVisual", DefaultParameterSetName = ParameterSetContext)]
[Alias("PdfVisual")]
[OutputType(typeof(PdfDocument))]
public sealed class AddOfficePdfVisualCommand : OfficeVisualCommandBase
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPipelineDocument = "PipelineDocument";

    /// <summary>ChartForgeX VisualArtifact, OfficeVisualSource, OfficeVisualConversionResult, or SVG file path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = ParameterSetContext)]
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPipelineDocument)]
    public object InputObject { get; set; } = null!;

    /// <summary>PDF document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetDocument)]
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetPipelineDocument)]
    public PdfDocument Document { get; set; } = null!;

    /// <summary>Horizontal alignment in PDF flow.</summary>
    [Parameter]
    public PdfAlign Align { get; set; } = PdfAlign.Left;

    /// <summary>Spacing before the visual in points.</summary>
    [Parameter]
    public double? SpacingBefore { get; set; }

    /// <summary>Spacing after the visual in points.</summary>
    [Parameter]
    public double? SpacingAfter { get; set; }

    /// <summary>Emit the updated PDF document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var visual = ResolveVisual(InputObject);
        var documentParameterSet = ParameterSetName == ParameterSetPipelineDocument
            ? ParameterSetPipelineDocument
            : ParameterSetDocument;
        var document = PdfCommandUtilities.ComposeContent(
            this,
            Document,
            ParameterSetName,
            documentParameterSet,
            content => content.AddVisualArtifact(visual, Align, SpacingBefore, SpacingAfter));
        if (PassThru.IsPresent && document != null)
        {
            WriteObject(document);
        }
    }
}
