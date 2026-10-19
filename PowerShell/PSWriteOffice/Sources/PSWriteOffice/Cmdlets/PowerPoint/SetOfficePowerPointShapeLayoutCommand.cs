using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Aligns, distributes, or arranges PowerPoint shapes using OfficeIMO layout helpers.</summary>
/// <example>
///   <summary>Align found shapes.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Find-OfficePowerPointShape -Slide $slide -Name 'Kpi.*' |
///     Set-OfficePowerPointShapeLayout -Align Top</code>
///   <para>Uses OfficeIMO.PowerPoint to align all matching shapes to the top edge of their selection bounds.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficePowerPointShapeLayout", DefaultParameterSetName = ParameterSetAlign)]
[Alias("PptShapeLayout", "PptArrange")]
[OutputType(typeof(PowerPointShape))]
public sealed class SetOfficePowerPointShapeLayoutCommand : PSCmdlet
{
    private const string ParameterSetAlign = "Align";
    private const string ParameterSetDistribute = "Distribute";
    private const string ParameterSetGrid = "Grid";

    private readonly List<object> _input = new();

    /// <summary>PowerPoint shapes or shape info records from Get-OfficePowerPointShape or Find-OfficePowerPointShape.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
    public object InputObject { get; set; } = null!;

    /// <summary>Slide that owns raw PowerPointShape inputs. Shape info records carry their own slide.</summary>
    [Parameter]
    public PowerPointSlide? Slide { get; set; }

    /// <summary>Alignment operation.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetAlign)]
    public PowerPointShapeAlignment Align { get; set; }

    /// <summary>Distribution operation.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetDistribute)]
    public PowerPointShapeDistribution Distribute { get; set; }

    /// <summary>Optional cross-axis alignment for even distribution.</summary>
    [Parameter(ParameterSetName = ParameterSetDistribute)]
    public PowerPointShapeAlignment? CrossAxisAlign { get; set; }

    /// <summary>Fixed spacing between distributed shapes in points.</summary>
    [Parameter(ParameterSetName = ParameterSetDistribute)]
    public double? SpacingPoints { get; set; }

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

    /// <summary>Horizontal grid gutter in points.</summary>
    [Parameter(ParameterSetName = ParameterSetGrid)]
    public double GutterXPoints { get; set; }

    /// <summary>Vertical grid gutter in points.</summary>
    [Parameter(ParameterSetName = ParameterSetGrid)]
    public double GutterYPoints { get; set; }

    /// <summary>Fill the grid column-by-column instead of row-by-row.</summary>
    [Parameter(ParameterSetName = ParameterSetGrid)]
    public PowerPointShapeGridFlow Flow { get; set; } = PowerPointShapeGridFlow.RowMajor;

    /// <summary>Keep each shape's current size when arranging in a grid.</summary>
    [Parameter(ParameterSetName = ParameterSetGrid)]
    public SwitchParameter NoResize { get; set; }

    /// <summary>Use the full slide bounds instead of the current selection bounds.</summary>
    [Parameter]
    public SwitchParameter ToSlide { get; set; }

    /// <summary>Use slide content bounds with the supplied margin in points.</summary>
    [Parameter]
    public double? MarginPoints { get; set; }

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
            case ParameterSetAlign:
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
        if (MarginPoints.HasValue)
        {
            slide.AlignShapesToSlideContentPoints(shapes, Align, MarginPoints.Value);
        }
        else if (ToSlide.IsPresent)
        {
            slide.AlignShapesToSlide(shapes, Align);
        }
        else
        {
            slide.AlignShapes(shapes, Align);
        }
    }

    private void InvokeDistribute(PowerPointSlide slide, IReadOnlyList<PowerPointShape> shapes)
    {
        if (SpacingPoints.HasValue)
        {
            if (MarginPoints.HasValue)
            {
                slide.DistributeShapesWithSpacingToSlideContentPoints(shapes, Distribute, SpacingPoints.Value, MarginPoints.Value, Center.IsPresent);
            }
            else
            {
                slide.DistributeShapesWithSpacingPoints(shapes, Distribute, SpacingPoints.Value, Center.IsPresent);
            }

            return;
        }

        if (MarginPoints.HasValue)
        {
            if (CrossAxisAlign.HasValue)
            {
                slide.DistributeShapesToSlideContentPoints(shapes, Distribute, MarginPoints.Value, CrossAxisAlign.Value);
            }
            else
            {
                slide.DistributeShapesToSlideContentPoints(shapes, Distribute, MarginPoints.Value);
            }
        }
        else if (ToSlide.IsPresent)
        {
            if (CrossAxisAlign.HasValue)
            {
                slide.DistributeShapesToSlide(shapes, Distribute, CrossAxisAlign.Value);
            }
            else
            {
                slide.DistributeShapesToSlide(shapes, Distribute);
            }
        }
        else if (CrossAxisAlign.HasValue)
        {
            slide.DistributeShapes(shapes, Distribute, CrossAxisAlign.Value);
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
            if (MarginPoints.HasValue)
            {
                slide.ArrangeShapesInGridAutoToSlideContentPoints(shapes, MarginPoints.Value, GutterXPoints, GutterYPoints, resizeToCell, Flow);
            }
            else if (ToSlide.IsPresent)
            {
                slide.ArrangeShapesInGridAutoToSlide(shapes, PowerPointUnits.FromPoints(GutterXPoints), PowerPointUnits.FromPoints(GutterYPoints), resizeToCell, Flow);
            }
            else
            {
                slide.ArrangeShapesInGridAuto(shapes, PowerPointUnits.FromPoints(GutterXPoints), PowerPointUnits.FromPoints(GutterYPoints), resizeToCell, Flow);
            }

            return;
        }

        if (MarginPoints.HasValue)
        {
            slide.ArrangeShapesInGridToSlideContentPoints(shapes, Columns!.Value, Rows!.Value, MarginPoints.Value, GutterXPoints, GutterYPoints, resizeToCell, Flow);
        }
        else if (ToSlide.IsPresent)
        {
            slide.ArrangeShapesInGridToSlide(shapes, Columns!.Value, Rows!.Value, PowerPointUnits.FromPoints(GutterXPoints), PowerPointUnits.FromPoints(GutterYPoints), resizeToCell, Flow);
        }
        else
        {
            slide.ArrangeShapesInGrid(shapes, Columns!.Value, Rows!.Value, PowerPointUnits.FromPoints(GutterXPoints), PowerPointUnits.FromPoints(GutterYPoints), resizeToCell, Flow);
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
        PowerPointSlide? slide = Slide;

        foreach (var item in _input)
        {
            switch (item)
            {
                case PowerPointShapeInfo info:
                    slide ??= info.Slide;
                    if (!ReferenceEquals(slide, info.Slide))
                    {
                        throw new PSArgumentException("All PowerPoint shapes must belong to the same slide.", nameof(InputObject));
                    }

                    shapes.Add(info.Shape);
                    break;
                case PowerPointShape shape:
                    shapes.Add(shape);
                    break;
                default:
                    throw new PSArgumentException("Input must contain PowerPointShape or PowerPointShapeInfo objects.", nameof(InputObject));
            }
        }

        if (slide == null)
        {
            throw new PSArgumentException("Provide -Slide when passing raw PowerPointShape objects.", nameof(Slide));
        }

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
