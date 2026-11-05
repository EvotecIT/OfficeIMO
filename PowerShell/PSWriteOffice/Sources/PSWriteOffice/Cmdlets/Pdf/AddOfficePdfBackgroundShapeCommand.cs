using OfficeIMO;
using OfficeIMO.Drawing;
using PSWriteOffice.Foundation;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Adds a decorative generated PDF page background shape or band.</summary>
/// <remarks>
/// Background shapes are intended for subtle page structure such as header bands, side bands, highlight panels, or decorative accents.
/// They are rendered behind generated content and should usually use restrained opacity values.
/// </remarks>
/// <example>
///   <summary>Add a header band and accent ellipse.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePdf -Path .\Report.pdf {
///     PdfBackgroundShape -Shape TopBand -Height 86 -FillColor '#DBEAFE' -FillOpacity 0.75
///     PdfBackgroundShape -Shape Ellipse -X 420 -Y 650 -Width 96 -Height 72 -FillColor '#99F6E4' -FillOpacity 0.35
///     PdfHeading 'Styled report'
///   }</code>
///   <para>Creates a polished generated page background without hand-drawing PDF primitives.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficePdfBackgroundShape", DefaultParameterSetName = ParameterSetContext)]
[Alias("PdfBackgroundShape") ]
public sealed class AddOfficePdfBackgroundShapeCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";

    /// <summary>Shape type to add.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetContext)]
    public OfficePdfBackgroundShapeType Shape { get; set; }

    /// <summary>Shape left coordinate in PDF points for explicit shapes.</summary>
    [Parameter]
    [OfficeLengthTransformation]
    public OfficeLength X { get; set; }

    /// <summary>Shape bottom coordinate in PDF points for explicit shapes.</summary>
    [Parameter]
    [OfficeLengthTransformation]
    public OfficeLength Y { get; set; }

    /// <summary>Shape width in PDF points, or band width for left/right bands.</summary>
    [Parameter]
    [OfficeLengthTransformation]
    public OfficeLength? Width { get; set; }

    /// <summary>Shape height in PDF points, or band height for top/bottom bands.</summary>
    [Parameter]
    [OfficeLengthTransformation]
    public OfficeLength? Height { get; set; }

    /// <summary>Rounded rectangle or band corner radius in PDF points.</summary>
    [Parameter]
    [OfficeLengthTransformation]
    public OfficeLength CornerRadius { get; set; }

    /// <summary>Horizontal inset for top/bottom bands in PDF points.</summary>
    [Parameter]
    [OfficeLengthTransformation]
    public OfficeLength InsetX { get; set; }

    /// <summary>Vertical inset for left/right bands in PDF points.</summary>
    [Parameter]
    [OfficeLengthTransformation]
    public OfficeLength InsetY { get; set; }

    /// <summary>Vertical offset for top/bottom bands in PDF points.</summary>
    [Parameter]
    [OfficeLengthTransformation]
    public OfficeLength OffsetY { get; set; }

    /// <summary>Horizontal offset for left/right bands in PDF points.</summary>
    [Parameter]
    [OfficeLengthTransformation]
    public OfficeLength OffsetX { get; set; }

    /// <summary>Fill color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? FillColor { get; set; }

    /// <summary>Stroke color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? StrokeColor { get; set; }

    /// <summary>Stroke width in PDF points.</summary>
    [Parameter]
    [OfficeLengthTransformation]
    public OfficeLength StrokeWidth { get; set; }

    /// <summary>Fill opacity from 0 to 1.</summary>
    [Parameter]
    [ValidateRange(0D, 1D)]
    public double? FillOpacity { get; set; }

    /// <summary>Stroke opacity from 0 to 1.</summary>
    [Parameter]
    [ValidateRange(0D, 1D)]
    public double? StrokeOpacity { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var fill = FillColor;
        var stroke = StrokeColor;
        PdfDslContext.Require(this).ConfigurePage( page =>
        {
            switch (Shape)
            {
                case OfficePdfBackgroundShapeType.Rectangle:
                    page.BackgroundRectangle(X, Y, RequireWidth(), RequireHeight(), fill, stroke, StrokeWidth, FillOpacity, StrokeOpacity);
                    break;
                case OfficePdfBackgroundShapeType.RoundedRectangle:
                    page.BackgroundRoundedRectangle(X, Y, RequireWidth(), RequireHeight(), CornerRadius, fill, stroke, StrokeWidth, FillOpacity, StrokeOpacity);
                    break;
                case OfficePdfBackgroundShapeType.Ellipse:
                    page.BackgroundEllipse(X, Y, RequireWidth(), RequireHeight(), fill, stroke, StrokeWidth, FillOpacity, StrokeOpacity);
                    break;
                case OfficePdfBackgroundShapeType.TopBand:
                    page.BackgroundTopBand(RequireHeight(), fill, InsetX, OffsetY, CornerRadius, stroke, StrokeWidth, FillOpacity, StrokeOpacity);
                    break;
                case OfficePdfBackgroundShapeType.BottomBand:
                    page.BackgroundBottomBand(RequireHeight(), fill, InsetX, OffsetY, CornerRadius, stroke, StrokeWidth, FillOpacity, StrokeOpacity);
                    break;
                case OfficePdfBackgroundShapeType.LeftBand:
                    page.BackgroundLeftBand(RequireWidth(), fill, InsetY, OffsetX, CornerRadius, stroke, StrokeWidth, FillOpacity, StrokeOpacity);
                    break;
                case OfficePdfBackgroundShapeType.RightBand:
                    page.BackgroundRightBand(RequireWidth(), fill, InsetY, OffsetX, CornerRadius, stroke, StrokeWidth, FillOpacity, StrokeOpacity);
                    break;
                default:
                    throw new PSArgumentOutOfRangeException(nameof(Shape), Shape, "Unsupported PDF background shape type.");
            }
        });

    }

    private OfficeLength RequireWidth()
    {
        return Width ?? throw new PSArgumentException("-Width is required for this background shape.", nameof(Width));
    }

    private OfficeLength RequireHeight()
    {
        return Height ?? throw new PSArgumentException("-Height is required for this background shape.", nameof(Height));
    }
}
