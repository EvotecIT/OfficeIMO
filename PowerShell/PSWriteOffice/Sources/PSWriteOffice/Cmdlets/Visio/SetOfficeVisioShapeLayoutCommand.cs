using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Visio;
using PSWriteOffice.Services.Visio;

namespace PSWriteOffice.Cmdlets.Visio;

/// <summary>Applies OfficeIMO Visio selection layout and layer operations to shapes.</summary>
/// <example>
///   <summary>Arrange keyed shapes in a horizontal stack.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>OfficeVisio -Path .\Layout.vsdx {
///     VisioRectangle -Key one -Text 'One'
///     VisioRectangle -Key two -Text 'Two'
///     VisioLayout -ShapeId one,two -HorizontalStack -HorizontalSpacing 0.4
/// }</code>
///   <para>Resolves keyed shapes and applies a reusable OfficeIMO Visio layout operation.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeVisioShapeLayout")]
[Alias("VisioLayout", "VisioArrange")]
[OutputType(typeof(VisioShape))]
public sealed class SetOfficeVisioShapeLayoutCommand : PSCmdlet
{
    private readonly List<object> _input = new();

    /// <summary>Shapes, shape selections, or shape keys/ids to arrange.</summary>
    [Parameter(Position = 0, ValueFromPipeline = true)]
    public object? InputObject { get; set; }

    /// <summary>Page that owns the shapes. Optional inside OfficeVisio/VisioPage DSL scopes.</summary>
    [Parameter]
    public VisioPage? Page { get; set; }

    /// <summary>Shape keys or ids to resolve on the target page.</summary>
    [Parameter]
    public string[]? ShapeId { get; set; }

    /// <summary>Add selected shapes to this layer.</summary>
    [Parameter]
    public string? Layer { get; set; }

    /// <summary>Horizontal alignment inside the selected shapes' bounds.</summary>
    [Parameter]
    public VisioHorizontalAlignment? AlignHorizontal { get; set; }

    /// <summary>Vertical alignment inside the selected shapes' bounds.</summary>
    [Parameter]
    public VisioVerticalAlignment? AlignVertical { get; set; }

    /// <summary>Distribute selected shapes along an axis.</summary>
    [Parameter]
    public VisioDistributionAxis? Distribute { get; set; }

    /// <summary>Lay out selected shapes as a grid.</summary>
    [Parameter]
    public SwitchParameter Grid { get; set; }

    /// <summary>Lay out selected shapes as a horizontal stack.</summary>
    [Parameter]
    public SwitchParameter HorizontalStack { get; set; }

    /// <summary>Lay out selected shapes as a vertical stack.</summary>
    [Parameter]
    public SwitchParameter VerticalStack { get; set; }

    /// <summary>Grid column count. Zero lets OfficeIMO choose a near-square grid.</summary>
    [Parameter]
    [ValidateRange(0, int.MaxValue)]
    public int Columns { get; set; }

    /// <summary>Horizontal spacing in inches for grid/stack layout.</summary>
    [Parameter]
    [ValidateRange(0, double.MaxValue)]
    public double HorizontalSpacing { get; set; } = 0.5D;

    /// <summary>Vertical spacing in inches for grid/stack layout.</summary>
    [Parameter]
    [ValidateRange(0, double.MaxValue)]
    public double VerticalSpacing { get; set; } = 0.5D;

    /// <summary>Use the first selected shape as the grid origin instead of preserving the selection top-left.</summary>
    [Parameter]
    public SwitchParameter PreserveFirstShapeCenter { get; set; }

    /// <summary>Do not reroute internal connectors during OfficeIMO relayout.</summary>
    [Parameter]
    public SwitchParameter NoRouteInternalConnectors { get; set; }

    /// <summary>Emit arranged shapes.</summary>
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
        var page = Page ?? VisioDslContext.Current?.CurrentPage;
        var shapes = ResolveShapes(page);
        var selection = new VisioShapeSelection(shapes, page);

        if (!string.IsNullOrWhiteSpace(Layer))
        {
            if (page != null)
            {
                foreach (var shape in shapes)
                {
                    page.AddToLayer(Layer!, shape);
                }
            }
            else
            {
                selection.Layer(Layer!);
            }
        }

        if (AlignHorizontal.HasValue)
        {
            selection.Align(AlignHorizontal.Value);
        }

        if (AlignVertical.HasValue)
        {
            selection.Align(AlignVertical.Value);
        }

        if (Distribute.HasValue)
        {
            selection.Distribute(Distribute.Value);
        }

        if (Grid.IsPresent)
        {
            selection.RelayoutAsGrid(new VisioSelectionLayoutOptions
            {
                Columns = Columns <= 0 ? null : Columns,
                HorizontalSpacing = HorizontalSpacing,
                VerticalSpacing = VerticalSpacing,
                PreserveTopLeft = !PreserveFirstShapeCenter.IsPresent,
                RouteInternalConnectors = !NoRouteInternalConnectors.IsPresent
            });
        }
        else if (HorizontalStack.IsPresent)
        {
            selection.RelayoutAsHorizontalStack(HorizontalSpacing, !NoRouteInternalConnectors.IsPresent);
        }
        else if (VerticalStack.IsPresent)
        {
            selection.RelayoutAsVerticalStack(VerticalSpacing, !NoRouteInternalConnectors.IsPresent);
        }

        if (PassThru.IsPresent)
        {
            WriteObject(shapes, enumerateCollection: true);
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

    private List<VisioShape> ResolveShapes(VisioPage? page)
    {
        var shapes = new List<VisioShape>();

        if (ShapeId != null)
        {
            foreach (var shapeId in ShapeId)
            {
                shapes.Add(ResolveShapeReference(page, shapeId));
            }
        }

        foreach (var item in _input)
        {
            switch (item)
            {
                case VisioShape shape:
                    shapes.Add(shape);
                    break;
                case string reference:
                    shapes.Add(ResolveShapeReference(page, reference));
                    break;
                default:
                    throw new PSArgumentException("Input must contain VisioShape objects, VisioShapeSelection objects, or shape keys/ids.", nameof(InputObject));
            }
        }

        return shapes.Distinct().ToList();
    }

    private VisioShape ResolveShapeReference(VisioPage? page, string reference)
    {
        if (string.IsNullOrWhiteSpace(reference))
        {
            throw new PSArgumentException("Shape reference cannot be empty.", nameof(ShapeId));
        }

        var effectivePage = page ?? VisioDslContext.Current?.CurrentPage;
        if (effectivePage == null)
        {
            throw new PSArgumentException("Provide -Page or run inside a VisioPage DSL scope when resolving shape references.", nameof(Page));
        }

        var context = VisioDslContext.Current;
        if (context != null)
        {
            return context.ResolveShape(effectivePage, reference);
        }

        var shape = effectivePage.AllShapes().FirstOrDefault(candidate =>
            string.Equals(candidate.Id, reference, System.StringComparison.OrdinalIgnoreCase) ||
            string.Equals(candidate.Name, reference, System.StringComparison.OrdinalIgnoreCase) ||
            string.Equals(candidate.NameU, reference, System.StringComparison.OrdinalIgnoreCase));

        return shape ?? throw new PSInvalidOperationException($"Visio shape '{reference}' was not found on the target Visio page.");
    }
}
