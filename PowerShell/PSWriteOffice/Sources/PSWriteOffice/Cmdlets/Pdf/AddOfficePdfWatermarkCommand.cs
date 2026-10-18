using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Adds a generated-document text watermark.</summary>
/// <example>
///   <summary>Add a draft watermark to every generated page.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePdf -Path .\Examples\Documents\PdfWatermark.pdf {
///     Add-OfficePdfWatermark -Text 'DRAFT' -FontSize 72 -Opacity 0.12 -RotationAngle -35 -Color '#64748B'
///     Add-OfficePdfHeading -Text 'Draft service review'
///     Add-OfficePdfParagraph -Text 'This copy is not final.'
/// }</code>
///   <para>Adds a text watermark while generating the PDF.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficePdfWatermark", DefaultParameterSetName = ParameterSetContext)]
[Alias("PdfWatermark")]
[OutputType(typeof(PdfDocument))]
public sealed class AddOfficePdfWatermarkCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>Compatibility parameter. Page composition is supported only inside New-OfficePdf.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public PdfDocument Document { get; set; } = null!;

    /// <summary>Watermark text.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Text { get; set; } = string.Empty;

    /// <summary>Watermark font size.</summary>
    [Parameter]
    public double? FontSize { get; set; }

    /// <summary>Watermark opacity, 0 through 1.</summary>
    [Parameter]
    public double? Opacity { get; set; }

    /// <summary>Watermark rotation angle.</summary>
    [Parameter]
    public double? RotationAngle { get; set; }

    /// <summary>Optional watermark color. Named colors and hexadecimal values are accepted.</summary>
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
        var document = PdfCommandUtilities.ComposePage(this, Document, ParameterSetName, ParameterSetDocument,
            page => page.Watermark(Text, FontSize, PdfCommandUtilities.ParseColor(Color), Opacity, RotationAngle));
        if (PassThru.IsPresent && document != null)
        {
            WriteObject(document);
        }
    }
}
