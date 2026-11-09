using System.Collections;
using System.Management.Automation;
using OfficeIMO;
using PSWriteOffice.Foundation;
using OfficeIMO.Visio;
using PSWriteOffice.Services.Visio;

namespace PSWriteOffice.Cmdlets.Visio;

/// <summary>Creates an OfficeIMO-authored Visio-native container around existing shapes.</summary>
/// <example>
///   <summary>Group related shapes in a container.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>OfficeVisio -Path .\Architecture.vsdx {
///     VisioRectangle -Key api -Text 'API' -X 2 -Y 4
///     VisioRectangle -Key worker -Text 'Worker' -X 4 -Y 4
///     VisioContainer -Key app -Text 'Application tier' -ShapeId api,worker -FillColor '#F8FAFC'
/// }</code>
///   <para>Creates a native Visio container around previously keyed shapes.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeVisioContainer", DefaultParameterSetName = "Page")]
[OutputType(typeof(VisioShape))]
[Alias("VisioContainer") ]
public sealed class AddOfficeVisioContainerCommand : OfficeVisioTextCmdlet {
    private readonly List<object> _input = new();

    /// <summary>Shapes, shape selections, or shape keys/ids to include in the container.</summary>
    [Parameter(Position = 0, ValueFromPipeline = true, ParameterSetName = "Page")]
    public object? InputObject { get; set; }

    /// <summary>Page that owns the member shapes. Optional inside OfficeVisio/VisioPage DSL scopes.</summary>
    [Parameter]
    public VisioPage? Page { get; set; }

    /// <summary>Shape keys or ids to include in the container.</summary>
    [Parameter]
    public string[]? ShapeId { get; set; }

    /// <summary>Container shape identifier.</summary>
    [Parameter(Mandatory = true)]
    public string Key { get; set; } = string.Empty;

    /// <summary>Container heading text.</summary>
    [Parameter]
    [AllowEmptyString]
    public string Text { get; set; } = string.Empty;

    /// <summary>Outer margin around member shapes as a physical length; a number means points.</summary>
    [Parameter]
    [OfficeLengthTransformation]
    public OfficeLength? Margin { get; set; }

    /// <summary>Additional heading height as a physical length; a number means points.</summary>
    [Parameter]
    [OfficeLengthTransformation]
    public OfficeLength? HeadingHeight { get; set; }

    /// <summary>Container fill color.</summary>
    [Parameter]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? FillColor { get; set; }

    /// <summary>Container line color.</summary>
    [Parameter]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? LineColor { get; set; }

    /// <summary>Container line weight as a physical length; a number means points.</summary>
    [Parameter]
    [OfficeLengthTransformation]
    public OfficeLength? LineWeight { get; set; }

    /// <summary>Native Visio container gallery identifier. Zero selects the default/no explicit gallery style.</summary>
    [Parameter]
    [ValidateRange(0, int.MaxValue)]
    public int? ContainerStyle { get; set; }

    /// <summary>Native Visio heading gallery identifier. Zero disables the heading style.</summary>
    [Parameter]
    [ValidateRange(0, int.MaxValue)]
    public int? HeadingStyle { get; set; }

    /// <summary>Lock the generated container.</summary>
    [Parameter]
    public SwitchParameter Locked { get; set; }

    /// <summary>Disable Visio container auto resize metadata.</summary>
    [Parameter]
    public SwitchParameter NoAutoResize { get; set; }

    /// <summary>Suppress Visio selection highlighting metadata.</summary>
    [Parameter]
    public SwitchParameter NoHighlight { get; set; }

    /// <summary>Suppress Visio container ribbon metadata.</summary>
    [Parameter]
    public SwitchParameter NoRibbon { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        AddInput(InputObject);
    }

    /// <inheritdoc />
    protected override void EndProcessing() {
        var page = Page ?? VisioDslContext.Current?.CurrentPage;
        if (page == null) {
            throw new PSArgumentException("Provide -Page or run inside a VisioPage DSL scope.", nameof(Page));
        }

        var shapes = ResolveShapes(page);
        if (shapes.Count == 0) {
            throw new PSArgumentException("At least one member shape is required.", nameof(InputObject));
        }

        var options = BuildOptions(page.DefaultUnit);
        var container = page.AddContainer(Key, Text, shapes, options);
        ApplyTextStyle(container);
        VisioDslContext.Current?.RegisterShape(page, Key, container);
        WritePassThru(container);
    }

    private VisioContainerOptions BuildOptions(VisioMeasurementUnit unit) {
        if (Margin?.ToPoints() < 0 || HeadingHeight?.ToPoints() < 0 || LineWeight?.ToPoints() < 0)
            throw new PSArgumentOutOfRangeException(nameof(Margin), Margin, "Container lengths must be non-negative.");
        var options = new VisioContainerOptions();

        if (Margin.HasValue) {
            options.Margin = ToPageUnits(Margin.Value, unit);
        }

        if (HeadingHeight.HasValue) {
            options.HeadingHeight = ToPageUnits(HeadingHeight.Value, unit);
        }

        if (FillColor.HasValue) {
            options.FillColor = FillColor.Value.Resolve(OfficeTheme.Default);
        }

        if (LineColor.HasValue) {
            options.LineColor = LineColor.Value.Resolve(OfficeTheme.Default);
        }

        if (LineWeight.HasValue) {
            options.LineWeight = LineWeight.Value.ToInches();
        }

        if (ContainerStyle.HasValue) {
            options.ContainerStyle = ContainerStyle.Value;
        }

        if (HeadingStyle.HasValue) {
            options.HeadingStyle = HeadingStyle.Value;
        }

        options.Locked = Locked.IsPresent;
        options.AutoResize = !NoAutoResize.IsPresent;
        options.NoHighlight = NoHighlight.IsPresent;
        options.NoRibbon = NoRibbon.IsPresent;
        return options;
    }

    private static double ToPageUnits(OfficeLength length, VisioMeasurementUnit unit) => unit switch {
        VisioMeasurementUnit.Centimeters => length.ToCentimeters(),
        VisioMeasurementUnit.Millimeters => length.ToMillimeters(),
        _ => length.ToInches()
    };

    private void AddInput(object? value) {
        if (value == null) {
            return;
        }

        if (value is PSObject psObject) {
            AddInput(psObject.BaseObject);
            return;
        }

        if (value is IEnumerable enumerable && value is not string) {
            foreach (var item in enumerable) {
                AddInput(item);
            }

            return;
        }

        _input.Add(value);
    }

    private List<VisioShape> ResolveShapes(VisioPage page) {
        var shapes = new List<VisioShape>();

        if (ShapeId != null) {
            foreach (var shapeId in ShapeId) {
                shapes.Add(ResolveShapeReference(page, shapeId));
            }
        }

        foreach (var item in _input) {
            switch (item) {
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

    private static VisioShape ResolveShapeReference(VisioPage page, string reference) {
        if (string.IsNullOrWhiteSpace(reference)) {
            throw new PSArgumentException("Shape reference cannot be empty.", nameof(ShapeId));
        }

        var context = VisioDslContext.Current;
        if (context != null) {
            return context.ResolveShape(page, reference);
        }

        var shape = page.AllShapes().FirstOrDefault(candidate =>
            string.Equals(candidate.Id, reference, System.StringComparison.OrdinalIgnoreCase) ||
            string.Equals(candidate.Name, reference, System.StringComparison.OrdinalIgnoreCase) ||
            string.Equals(candidate.NameU, reference, System.StringComparison.OrdinalIgnoreCase));

        return shape ?? throw new PSInvalidOperationException($"Visio shape '{reference}' was not found on the target Visio page.");
    }
}
