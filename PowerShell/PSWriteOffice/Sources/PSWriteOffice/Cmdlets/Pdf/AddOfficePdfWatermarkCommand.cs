using OfficeIMO;
using OfficeIMO.Drawing;
using PSWriteOffice.Foundation;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Adds a generated-document text watermark.</summary>
/// <example>
///   <summary>Add a draft watermark to every generated page.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePdf -Path .\Examples\Documents\PdfWatermark.pdf {
///     Add-OfficePdfWatermark -Text 'DRAFT' -Size 72 -Opacity 0.12 -Rotation -35 -Color '#64748B'
///     Add-OfficePdfHeading -Text 'Draft service review'
///     Add-OfficePdfParagraph -Text 'This copy is not final.'
/// }</code>
///   <para>Adds a text watermark while generating the PDF.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficePdfWatermark", DefaultParameterSetName = ParameterSetContext)]
[Alias("PdfWatermark") ]
public sealed class AddOfficePdfWatermarkCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";

    /// <summary>Watermark text.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetContext)]
    public string Text { get; set; } = string.Empty;

    /// <summary>Watermark font size.</summary>
    [Parameter]
    [OfficeLengthTransformation]
    public OfficeLength? Size { get; set; }

    /// <summary>Watermark opacity, 0 through 1.</summary>
    [Parameter]
    public double? Opacity { get; set; }

    /// <summary>Watermark rotation angle.</summary>
    [Parameter]
    public double? Rotation { get; set; }

    /// <summary>Optional watermark color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? Color { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        PdfDslContext.Require(this).ConfigurePage(
            page => page.Watermark(Text, Size, Color, Opacity, Rotation));
    }
}
