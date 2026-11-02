using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Aligns, distributes, or arranges PowerPoint shapes using OfficeIMO layout helpers.</summary>
/// <example>
///   <summary>Align found shapes.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficePowerPointShape -Slide $slide -Name 'Kpi.*' |
///     Set-OfficePowerPointShapeLayout -Alignment Top</code>
///   <para>Uses OfficeIMO.PowerPoint to align all matching shapes to the top edge of their selection bounds.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficePowerPointShapeLayout", DefaultParameterSetName = ParameterSetAlignment)]
[Alias("PowerPointShapeLayout")]
[OutputType(typeof(PowerPointShape))]
public sealed class SetOfficePowerPointShapeLayoutCommand : PSCmdlet
{
    private const string ParameterSetAlignment = "Alignment";
    private const string ParameterSetDistribute = "Distribute";
    private const string ParameterSetGrid = "Grid";

    private readonly List<object> _input = new();

    /// <summary>PowerPoint shapes or shape info records from Get-OfficePowerPointShape.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
    public object[] InputObject { get; set; } = Array.Empty<object>();

    /// <summary>Native slide or one-based slide number; omission uses the supplied shapes' owning slide.</summary>
    [Parameter, PowerPointSlideTransformation]
    public PowerPointSlideSelector? Slide { get; set; }

    /// <summary>Alignment operation.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetAlignment)]
    public PowerPointShapeAlignment Alignment { get; set; }

    /// <summary>Distribution operation.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetDistribute)]
    public PowerPointShapeDistribution Distribute { get; set; }

    /// <summary>Optional cross-axis alignment for even distribution.</summary>
    [Parameter(ParameterSetName = ParameterSetDistribute)]
    public PowerPointShapeAlignment? CrossAxisAlignment { get; set; }

    /// <summary>Fixed spacing between distributed shapes; bare numbers mean points.</summary>
    [Parameter(ParameterSetName = ParameterSetDistribute)]
    [OfficeLengthTransformation]
    public OfficeLength? Spacing { get; set; }

    /// <summary>Arrange shapes in a grid.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetGrid)]
    public SwitchParameter Grid { get; set; }

    /// <summary>Grid column count. Omit with AutoGrid.</summary>
    [Parameter(ParameterSetName = ParameterSetGrid)]
    public int? Columns { get; set; }

    /// <summary>Grid row count. Omit with AutoGrid.</summary>
    [Parameter(ParameterSetName = ParameterSetGrid)]
    public int? Rows { get; set; }

    /// <summary>Let OfficeIMO choose the grid dimensions.</summary>
    [Parameter(ParameterSetName = ParameterSetGrid)]
    public SwitchParameter AutoGrid { get; set; }

    /// <summary>Horizontal grid gutter; bare numbers mean points.</summary>
    [Parameter(ParameterSetName = ParameterSetGrid)]
    [OfficeLengthTransformation]
    public OfficeLength GutterX { get; set; }

    /// <summary>Vertical grid gutter; bare numbers mean points.</summary>
    [Parameter(ParameterSetName = ParameterSetGrid)]
    [OfficeLengthTransformation]
    public OfficeLength GutterY { get; set; }

    /// <summary>Fill the grid column-by-column instead of row-by-row.</summary>
    [Parameter(ParameterSetName = ParameterSetGrid)]
    public PowerPointShapeGridFlow Flow { get; set; } = PowerPointShapeGridFlow.RowMajor;

    /// <summary>Keep each shape's current size when arranging in a grid.</summary>
    [Parameter(ParameterSetName = ParameterSetGrid)]
    public SwitchParameter NoResize { get; set; }

    /// <summary>Use the full slide bounds instead of the current selection bounds.</summary>
    [Parameter]
    public SwitchParameter ToSlide { get; set; }

    /// <summary>Use slide content bounds with the supplied margin; bare numbers mean points.</summary>
    [Parameter]
    [OfficeLengthTransformation]
    public OfficeLength? Margin { get; set; }

    /// <summary>Center a fixed-spacing distribution within its bounds.</summary>
    [Parameter(ParameterSetName = ParameterSetDistribute)]
    public SwitchParameter Center { get; set; }

    /// <summary>Emit the arranged shapes.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        AddInput(InputObject);
    }

    /// <inheritdoc />
    protected override void EndProcessing()
    {
        var resolved = ResolveShapes();
        if (resolved.Shapes.Count == 0)
        {
            throw new PSArgumentException("Provide at least one PowerPoint shape.", nameof(InputObject));
        }

        switch (ParameterSetName)
        {
            case ParameterSetAlignment:
                InvokeAlign(resolved.Slide, resolved.Shapes);
                break;
            case ParameterSetDistribute:
                InvokeDistribute(resolved.Slide, resolved.Shapes);
                break;
            default:
                InvokeGrid(resolved.Slide, resolved.Shapes);
                break;
        }

        if (PassThru.IsPresent)
        {
            WriteObject(resolved.Shapes, enumerateCollection: true);
        }
    }

    private void InvokeAlign(PowerPointSlide slide, IReadOnlyList<PowerPointShape> shapes)
    {
        if (Margin.HasValue)
        {
            slide.AlignShapesToSlideContent(shapes, Alignment, Margin.Value);
        }
        else if (ToSlide.IsPresent)
        {
            slide.AlignShapesToSlide(shapes, Alignment);
        }
        else
        {
            slide.AlignShapes(shapes, Alignment);
        }
    }

    private void InvokeDistribute(PowerPointSlide slide, IReadOnlyList<PowerPointShape> shapes)
    {
        if (Spacing.HasValue)
        {
            var options = new PowerPointShapeSpacingOptions {
                Spacing = Spacing.Value,
                CrossAxisAlignment = CrossAxisAlignment,
                Alignment = Center.IsPresent ?
                    Distribute == PowerPointShapeDistribution.Horizontal ? PowerPointShapeAlignment.Center : PowerPointShapeAlignment.Middle : null
            };
            if (Margin.HasValue) slide.DistributeShapesWithSpacingToSlideContent(shapes, Distribute, Margin.Value, options);
            else if (ToSlide.IsPresent) slide.DistributeShapesWithSpacingToSlide(shapes, Distribute, options);
            else slide.DistributeShapesWithSpacing(shapes, Distribute, options);
            return;
        }

        if (Margin.HasValue)
        {
            if (CrossAxisAlignment.HasValue)
            {
                slide.DistributeShapesToSlideContent(shapes, Distribute, Margin.Value, CrossAxisAlignment.Value);
            }
            else
            {
                slide.DistributeShapesToSlideContent(shapes, Distribute, Margin.Value);
            }
        }
        else if (ToSlide.IsPresent)
        {
            if (CrossAxisAlignment.HasValue)
            {
                slide.DistributeShapesToSlide(shapes, Distribute, CrossAxisAlignment.Value);
            }
            else
            {
                slide.DistributeShapesToSlide(shapes, Distribute);
            }
        }
        else if (CrossAxisAlignment.HasValue)
        {
            slide.DistributeShapes(shapes, Distribute, CrossAxisAlignment.Value);
        }
        else
        {
            slide.DistributeShapes(shapes, Distribute);
        }
    }

    private void InvokeGrid(PowerPointSlide slide, IReadOnlyList<PowerPointShape> shapes)
    {
        if (!AutoGrid.IsPresent && (!Columns.HasValue || !Rows.HasValue))
        {
            throw new PSArgumentException("Specify -Columns and -Rows, or use -AutoGrid.", nameof(Columns));
        }

        var resizeToCell = !NoResize.IsPresent;
        if (AutoGrid.IsPresent)
        {
            if (Margin.HasValue)
            {
                slide.ArrangeShapesInGridAutoToSlideContent(shapes, Margin.Value, GutterX, GutterY, resizeToCell, Flow);
            }
            else if (ToSlide.IsPresent)
            {
                slide.ArrangeShapesInGridAutoToSlide(shapes, GutterX, GutterY, resizeToCell, Flow);
            }
            else
            {
                slide.ArrangeShapesInGridAuto(shapes, GutterX, GutterY, resizeToCell, Flow);
            }

            return;
        }

        if (Margin.HasValue)
        {
            slide.ArrangeShapesInGridToSlideContent(shapes, Columns!.Value, Rows!.Value, Margin.Value, GutterX, GutterY, resizeToCell, Flow);
        }
        else if (ToSlide.IsPresent)
        {
            slide.ArrangeShapesInGridToSlide(shapes, Columns!.Value, Rows!.Value, GutterX, GutterY, resizeToCell, Flow);
        }
        else
        {
            slide.ArrangeShapesInGrid(shapes, Columns!.Value, Rows!.Value, GutterX, GutterY, resizeToCell, Flow);
        }
    }

    private void AddInput(object? value)
    {
        if (value == null)
        {
            return;
        }

        if (value is PSObject psObject)
        {
            AddInput(psObject.BaseObject);
            return;
        }

        if (value is IEnumerable enumerable && value is not string)
        {
            foreach (var item in enumerable)
            {
                AddInput(item);
            }

            return;
        }

        _input.Add(value);
    }

    private ResolvedShapeSet ResolveShapes()
    {
        var shapes = new List<PowerPointShape>();
        PowerPointSlide? slide = null;
        foreach (var item in _input)
        {
            var shape = item switch
            {
                PowerPointShapeInfo info => info.Shape,
                PowerPointShape native => native,
                _ => throw new PSArgumentException("Input must contain PowerPointShape or PowerPointShapeInfo objects.", nameof(InputObject))
            };
            var owner = shape.OwnerSlide ?? throw new PSArgumentException("The shape has no owning slide.", nameof(InputObject));
            slide ??= owner;
            if (!ReferenceEquals(slide, owner))
                throw new PSArgumentException("All PowerPoint shapes must belong to the same slide.", nameof(InputObject));
            shapes.Add(shape);
        }
        if (slide == null) throw new PSArgumentException("Provide at least one PowerPoint shape.", nameof(InputObject));
        if (Slide != null && !ReferenceEquals(Slide.Resolve(slide.OwnerDocument), slide))
            throw new PSArgumentException("The selected slide does not own the supplied shapes.", nameof(Slide));

        return new ResolvedShapeSet(slide, shapes);
    }

    private sealed class ResolvedShapeSet
    {
        public ResolvedShapeSet(PowerPointSlide slide, IReadOnlyList<PowerPointShape> shapes)
        {
            Slide = slide;
            Shapes = shapes;
        }

        public PowerPointSlide Slide { get; }

        public IReadOnlyList<PowerPointShape> Shapes { get; }
    }
}
