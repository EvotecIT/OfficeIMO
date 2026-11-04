using System;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Foundation;

/// <summary>Shared native slide selection and physical placement for PowerPoint text content.</summary>
public abstract class PowerPointTextPlacementCmdlet : OfficeCmdlet {
    /// <summary>Gets or sets a native slide or one-based position; omission uses the current slide block.</summary>
    [Parameter(ValueFromPipeline = true), PowerPointSlideTransformation]
    public PowerPointSlideSelector? Slide { get; set; }
    /// <summary>Gets or sets the left coordinate; numbers mean points.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? X { get; set; }
    /// <summary>Gets or sets the top coordinate; numbers mean points.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? Y { get; set; }
    /// <summary>Gets or sets the width; omission uses the native default for fixed placement.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? Width { get; set; }
    /// <summary>Gets or sets the height; omission uses the native default for fixed placement.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? Height { get; set; }
    /// <summary>Gets or sets the one-based theme grid column; explicit X overrides its position.</summary>
    [Parameter, ValidateRange(1, 12)]
    public int? Column { get; set; }
    /// <summary>Gets or sets the grid column count; omission with Column uses twelve. Explicit Width overrides it.</summary>
    [Parameter, ValidateRange(1, 12)]
    public int? Span { get; set; }

    /// <summary>Explicit selection or geometry opts out of automatic content flow.</summary>
    private protected bool HasFixedPlacement => Slide != null || Column.HasValue || Span.HasValue ||
        PowerPointManualPlacement.HasExplicitBounds(MyInvocation.BoundParameters.Keys);

    /// <summary>Creates a native frame after validating lengths, slide selection and grid geometry.</summary>
    private protected (PowerPointTextBox Box, PowerPointContentScope? Scope) CreatePlacedTextBox(string text) {
        if (Width is OfficeLength w && w <= OfficeLength.Zero || Height is OfficeLength h && h <= OfficeLength.Zero)
            throw new ArgumentException("Text-box width and height must be positive.");
        var target = PowerPointManualPlacement.Resolve(Slide);
        var horizontal = PowerPointManualPlacement.Horizontal(target.Slide, X, Width, Column, Span);
        return (target.Slide.AddTextBox(text, horizontal.X, Y, horizontal.Width, Height), target.Scope);
    }
}
