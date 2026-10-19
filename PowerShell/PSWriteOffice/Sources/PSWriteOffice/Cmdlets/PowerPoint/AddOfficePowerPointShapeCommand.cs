using System;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Adds a basic shape to a slide.</summary>
/// <para>Creates an auto shape at the requested coordinates and applies optional fill and outline styling.</para>
/// <example>
///   <summary>Create a rectangle highlight.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePowerPoint -Path .\Examples\Documents\PowerPointShape.pptx {
///     $slide = Add-OfficePowerPointSlide -Layout 1 -PassThru
///     Add-OfficePowerPointShape -Slide $slide -ShapeType Rectangle -X 60 -Y 120 -Width 220 -Height 90 -FillColor '#DDEEFF' -OutlineColor '#2563EB' -OutlineWidth 1
///     Add-OfficePowerPointTextBox -Slide $slide -Text 'Highlighted status' -X 80 -Y 145 -Width 180 -Height 32
/// }</code>
///   <para>Creates a styled rectangle and overlays a text box.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficePowerPointShape")]
[Alias("PptShape")]
public sealed class AddOfficePowerPointShapeCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    /// <summary>Target slide that will receive the shape (optional inside DSL).</summary>
    [Parameter(ValueFromPipeline = true)]
    public PowerPointSlide? Slide { get; set; }

    /// <summary>Shape geometry preset name (e.g., Rectangle, Ellipse, Line).</summary>
    [Parameter]
    public OfficePresetShapeType ShapeType { get; set; } = OfficePresetShapeType.Rectangle;

    /// <summary>Left offset (in points) from the slide origin.</summary>
    [Parameter]
    public double X { get; set; } = 50;

    /// <summary>Top offset (in points) from the slide origin.</summary>
    [Parameter]
    public double Y { get; set; } = 50;

    /// <summary>Shape width in points.</summary>
    [Parameter]
    public double Width { get; set; } = 200;

    /// <summary>Shape height in points.</summary>
    [Parameter]
    public double Height { get; set; } = 100;

    /// <summary>Optional name assigned to the shape.</summary>
    [Parameter]
    public string? Name { get; set; }

    /// <summary>Fill color (hex or named color).</summary>
    [Parameter]
    [OfficeColorArgumentTransformation]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? FillColor { get; set; }

    /// <summary>Outline color (hex or named color).</summary>
    [Parameter]
    [OfficeColorArgumentTransformation]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? OutlineColor { get; set; }

    /// <summary>Outline width in points.</summary>
    [Parameter]
    public double? OutlineWidth { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        try {
            if (Width <= 0) {
                throw new ArgumentOutOfRangeException(nameof(Width), "Width must be greater than 0.");
            }

            if (Height <= 0) {
                throw new ArgumentOutOfRangeException(nameof(Height), "Height must be greater than 0.");
            }

            if (OutlineWidth is < 0) {
                throw new ArgumentOutOfRangeException(nameof(OutlineWidth), "OutlineWidth cannot be negative.");
            }

            var slide = Slide ?? PowerPointDslContext.Require(this).RequireSlide();
            var shape = slide.AddShapePoints(ShapeType, X, Y, Width, Height, Name);

            var fill = NormalizeColor(FillColor);
            if (fill != null) {
                shape.FillColor = fill;
            }

            var outline = NormalizeColor(OutlineColor);
            if (outline != null) {
                shape.OutlineColor = outline;
            }

            if (OutlineWidth.HasValue) {
                shape.OutlineWidthPoints = OutlineWidth.Value;
            }

            WritePassThru(shape);
        } catch (Exception ex) {
            WriteError(new ErrorRecord(ex, "PowerPointAddShapeFailed", ErrorCategory.InvalidOperation, Slide));
        }
    }

    private static string? NormalizeColor(string? color) {
        if (string.IsNullOrWhiteSpace(color)) {
            return null;
        }

        return OfficeColor.Parse(color!).ToRgbHex().ToLowerInvariant();
    }

}
