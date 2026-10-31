using System;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Adds a native preset shape at physical coordinates or on the theme grid.</summary>
/// <example><summary>Draw a named status light.</summary>
/// <code>PowerPointSlide -Title 'Service health' {
///     PowerPointShape -Type Ellipse -Name 'Healthy' -FillColor Success -X 2cm -Y 4cm -Width 2cm -Height 2cm
/// }</code></example>
[Cmdlet(VerbsCommon.Add, "OfficePowerPointShape")]
[Alias("PowerPointShape")]
[OutputType(typeof(PowerPointAutoShape))]
public sealed class AddOfficePowerPointShapeCommand : OfficeCmdlet {
    /// <summary>Gets or sets a native slide or one-based position; omission uses the current slide block.</summary>
    [Parameter(ValueFromPipeline = true), PowerPointSlideTransformation]
    public PowerPointSlideSelector? Slide { get; set; }
    /// <summary>Gets or sets the native preset geometry.</summary>
    [Parameter]
    public OfficePresetShapeType Type { get; set; } = OfficePresetShapeType.Rectangle;
    /// <summary>Gets or sets the native shape name.</summary>
    [Parameter]
    public string? Name { get; set; }
    /// <summary>Gets or sets the fill color or theme role.</summary>
    [Parameter, OfficeColorTransformation]
    public OfficeColor? FillColor { get; set; }
    /// <summary>Gets or sets the outline color or theme role.</summary>
    [Parameter, OfficeColorTransformation]
    public OfficeColor? OutlineColor { get; set; }
    /// <summary>Gets or sets the outline thickness; numbers mean points.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? OutlineWidth { get; set; }
    /// <summary>Gets or sets the left coordinate; numbers mean points.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? X { get; set; }
    /// <summary>Gets or sets the top coordinate; numbers mean points.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? Y { get; set; }
    /// <summary>Gets or sets the width; omission uses the native default.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? Width { get; set; }
    /// <summary>Gets or sets the height; omission uses the native default.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? Height { get; set; }
    /// <summary>Gets or sets the one-based theme grid column; explicit X overrides its position.</summary>
    [Parameter, ValidateRange(1, 12)]
    public int? Column { get; set; }
    /// <summary>Gets or sets the number of grid columns; the default is twelve, and explicit Width overrides it.</summary>
    [Parameter, ValidateRange(1, 12)]
    public int? Span { get; set; }
    /// <summary>Gets or sets whether the live native shape is returned.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        PowerPointAutoShape? shape = null;
        try {
            if (Width is OfficeLength w && w <= OfficeLength.Zero || Height is OfficeLength h && h <= OfficeLength.Zero)
                throw new ArgumentException("Shape width and height must be positive.");
            if (OutlineWidth.HasValue) PowerPointShape.ToDrawingLineWidth(OutlineWidth.Value, nameof(OutlineWidth));
            var target = PowerPointManualPlacement.Resolve(Slide);
            var document = target.Slide.OwnerDocument ?? throw new ArgumentException("The slide has no owning document.");
            var context = new PowerPointLayoutContext(document, PowerPointLayoutTheme.From(target.Slide.AuthoringTheme));
            string? fill = FillColor.HasValue ? context.ResolveColor(FillColor) : null;
            string? outline = OutlineColor.HasValue ? context.ResolveColor(OutlineColor) : null;
            var horizontal = PowerPointManualPlacement.Horizontal(target.Slide, X, Width, Column, Span);
            shape = target.Slide.AddShape(Type, horizontal.X, Y, horizontal.Width, Height, Name);
            if (fill != null) shape.FillColor = fill;
            if (outline != null) shape.OutlineColor = outline;
            if (OutlineWidth.HasValue) shape.OutlineWidth = OutlineWidth;
            target.Scope?.TrackFixedFigure(shape);
            if (PassThru) WriteObject(shape);
        } catch (Exception error) {
            shape?.Remove();
            if (error is PipelineStoppedException or ActionPreferenceStopException) throw;
            if (PowerPointContentScope.Current != null)
                ThrowTerminatingError(OfficeError.Create(error, "powerpoint-shape-create-failed", Slide, ErrorCategory.InvalidArgument));
            if (error is ArgumentException argument) InvalidArgument(argument, Slide);
            else ItemError(error, "powerpoint-shape-create-failed", Slide);
        }
    }
}
