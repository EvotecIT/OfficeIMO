using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Adds an image to a PDF document.</summary>
/// <example>
///   <summary>Add a logo or chart image to a PDF.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$image = '.\Tests\Assets\CellImage.png'
/// New-OfficePdf -Path .\Examples\Documents\PdfImage.pdf {
///     Add-OfficePdfHeading -Text 'Visual evidence'
///     Add-OfficePdfImage -Path $image -Width 120 -Height 80 -Align Center -AlternativeText 'Sample chart image'
/// }</code>
///   <para>Embeds an image with alternative text in the generated PDF.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficePdfImage", DefaultParameterSetName = ParameterSetContext)]
[Alias("PdfImage")]
[OutputType(typeof(PdfDocument))]
public sealed class AddOfficePdfImageCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>PDF document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public PdfDocument Document { get; set; } = null!;

    /// <summary>Image path.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Rendered image width in PDF points.</summary>
    [Parameter(Mandatory = true)]
    public double Width { get; set; }

    /// <summary>Rendered image height in PDF points.</summary>
    [Parameter(Mandatory = true)]
    public double Height { get; set; }

    /// <summary>Image alignment.</summary>
    [Parameter]
    public PdfAlign Align { get; set; } = PdfAlign.Left;

    /// <summary>Alternative text for meaningful images.</summary>
    [Parameter]
    public string? AlternativeText { get; set; }

    /// <summary>Emit the updated document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var imagePath = PdfCommandUtilities.ResolvePath(this, Path);
        var imageBytes = PdfImageDocumentSource.FromFile(imagePath).GetBytes();
        var document = PdfCommandUtilities.ComposeContent(this, Document, ParameterSetName, ParameterSetDocument,
            content => content.Image(imageBytes, Width, Height, Align, null, null, null, null, null, null, null, AlternativeText));
        if (PassThru.IsPresent && document != null)
        {
            WriteObject(document);
        }
    }
}
