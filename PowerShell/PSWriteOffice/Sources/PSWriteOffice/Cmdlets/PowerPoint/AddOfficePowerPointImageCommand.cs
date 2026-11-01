using System;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.ChartForgeX;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.PowerPoint;
using PSWriteOffice.Services.Visuals;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Adds a native picture from a local image, SVG or reusable visual, using the slide layout or explicit lengths.</summary>
/// <para>Coordinates use points or unit suffixes. Explicit coordinates, dimensions, grid columns or Slide select fixed placement. Otherwise the image flows through the shared slide layout. FocalPoint controls cover crops.</para>
/// <example><summary>Place a local image on the slide grid.</summary>
/// <code>PowerPointSlide -Title 'Evidence' {
///     PowerPointImage -Path ./evidence.png -Column 2 -Span 5 -Y 3cm -Height 5cm -AlternativeText 'Deployment health'
/// }</code></example>
/// <example><summary>Reuse a converted visual without converting it again.</summary>
/// <code>$visual | PowerPointImage -Slide $slide -X 36 -Y 54 -PassThru</code></example>
[Cmdlet(VerbsCommon.Add, "OfficePowerPointImage", DefaultParameterSetName = "Path")]
[Alias("PowerPointImage")]
[OutputType(typeof(PowerPointPicture))]
public sealed class AddOfficePowerPointImageCommand : OfficeVisualCommandBase {
    /// <summary>Accessible description for the visual or placed image.</summary>
    [Parameter]
    public string? AlternativeText { get; set; }

    /// <inheritdoc />
    protected override string? VisualAlternativeText => AlternativeText;

    /// <summary>Native slide or a one-based position in the current document; omitted inside PowerPointSlide.</summary>
    [Parameter, PowerPointSlideTransformation]
    public PowerPointSlideSelector? Slide { get; set; }

    /// <summary>Literal local image path.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Path"), ValidateNotNullOrEmpty]
    public string? Path { get; set; }

    /// <summary>Visual artifact, OfficeVisualSource, prepared conversion or SVG path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Visual"), ValidateNotNull]
    public object? InputObject { get; set; }

    /// <summary>Left position; numbers mean points.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? X { get; set; }

    /// <summary>Top position; numbers mean points.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? Y { get; set; }

    /// <summary>One-based starting column in the native twelve-column slide grid.</summary>
    [Parameter, ValidateRange(1, 12)]
    public int? Column { get; set; }

    /// <summary>Number of grid columns occupied by the frame; defaults to twelve with Column.</summary>
    [Parameter, ValidateRange(1, 12)]
    public int? Span { get; set; }

    /// <summary>Center, Top, Bottom, Left, Right or fractional x,y coordinates kept in view by a cover crop.</summary>
    [Parameter, PowerPointFocalPointTransformation]
    public PowerPointFocalPoint FocalPoint { get; set; } = new(0.5, 0.5);

    /// <summary>Return the native editable picture.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        PowerPointPicture? picture = null;
        var scope = PowerPointContentScope.Current;
        try {
            if (Width.HasValue && Width.Value <= OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(Width));
            if (Height.HasValue && Height.Value <= OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(Height));
            var target = PowerPointManualPlacement.Resolve(Slide);
            var horizontal = PowerPointManualPlacement.Horizontal(target.Slide, X, Width, Column, Span);
            OfficeVisualConversionResult? visual = ParameterSetName == "Visual" ? ResolveVisual(InputObject!, nativePlacement: true) : null;
            string? path = Path == null ? null : OfficePath.ResolveInput(SessionState, Path, literal: true)[0];
            OfficeLength width = horizontal.Width ?? OfficeLength.Points(visual?.WidthPoints ?? 200);
            OfficeLength height = Height ?? OfficeLength.Points(visual?.HeightPoints ?? 150);
            OfficeLength left = horizontal.X ?? OfficeLength.Zero, top = Y ?? OfficeLength.Zero;
            picture = visual == null ? target.Slide.AddImage(path!, left, top, width, height)
                : target.Slide.AddVisualArtifact(visual, new PowerPointLayoutBox(left, top, width, height));
            if (AlternativeText != null) picture.Description = AlternativeText;
            if (Title != null) picture.Title = Title;
            bool fixedPlacement = Slide != null || Column.HasValue || Span.HasValue
                || PowerPointManualPlacement.HasExplicitBounds(MyInvocation.BoundParameters.Keys);
            if (fixedPlacement && MyInvocation.BoundParameters.ContainsKey(nameof(FocalPoint))) {
                picture.Left = left; picture.Top = top; picture.Width = width; picture.Height = height;
                if (!picture.TryGetPixelSize(out int pixelWidth, out int pixelHeight))
                    throw new ArgumentException("The image dimensions are required for focal-point cropping.");
                picture.CropToFill(pixelWidth, pixelHeight, FocalPoint.X, FocalPoint.Y);
            }
            if (target.Scope != null) {
                if (fixedPlacement) target.Scope.TrackFixedFigure(picture);
                else target.Scope.AddFigure(new PowerPointLayoutImage(picture) {
                    FocalX = FocalPoint.X, FocalY = FocalPoint.Y
                }, picture);
            }
            if (PassThru) WriteObject(picture);
        } catch (Exception error) {
            picture?.Remove();
            if (error is PipelineStoppedException or ActionPreferenceStopException) throw;
            var record = OfficeError.Create(error, "invalid-argument", InputObject ?? (object?)Path, ErrorCategory.InvalidArgument);
            if (scope != null) ThrowTerminatingError(record);
            else WriteError(record);
        }
    }
}
