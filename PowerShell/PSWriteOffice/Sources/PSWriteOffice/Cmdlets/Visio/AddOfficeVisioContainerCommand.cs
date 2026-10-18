using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Drawing;
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
///     VisioContainer -Id app -Text 'Application tier' -ShapeId api,worker -FillColor '#F8FAFC'
/// }</code>
///   <para>Creates a native Visio container around previously keyed shapes.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeVisioContainer")]
[Alias("VisioContainer")]
[OutputType(typeof(VisioShape))]
public sealed class AddOfficeVisioContainerCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    private readonly List<object> _input = new();

    /// <summary>Shapes, shape selections, or shape keys/ids to include in the container.</summary>
    [Parameter(Position = 0, ValueFromPipeline = true)]
    public object? InputObject { get; set; }

    /// <summary>Page that owns the member shapes. Optional inside OfficeVisio/VisioPage DSL scopes.</summary>
    [Parameter]
    public VisioPage? Page { get; set; }

    /// <summary>Shape keys or ids to include in the container.</summary>
    [Parameter]
    public string[]? ShapeId { get; set; }

    /// <summary>Container shape identifier.</summary>
    [Parameter(Mandatory = true)]
    public string Id { get; set; } = string.Empty;

    /// <summary>Container heading text.</summary>
    [Parameter]
    [AllowEmptyString]
    public string Text { get; set; } = string.Empty;

    /// <summary>Outer margin around member shapes in page units.</summary>
    [Parameter]
    [ValidateRange(0, double.MaxValue)]
    public double? Margin { get; set; }

    /// <summary>Additional heading height in page units.</summary>
    [Parameter]
    [ValidateRange(0, double.MaxValue)]
    public double? HeadingHeight { get; set; }

    /// <summary>Container fill color.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? FillColor { get; set; }

    /// <summary>Container line color.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? LineColor { get; set; }

    /// <summary>Container line weight in inches.</summary>
    [Parameter]
    [ValidateRange(0, double.MaxValue)]
    public double? LineWeight { get; set; }

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

        var options = BuildOptions();
        var container = page.AddContainer(Id, Text, shapes, options);
        VisioDslContext.Current?.RegisterShape(page, Id, container);
        WritePassThru(container);
    }

    private VisioContainerOptions BuildOptions() {
        var options = new VisioContainerOptions();

        if (Margin.HasValue) {
            options.Margin = Margin.Value;
        }

        if (HeadingHeight.HasValue) {
            options.HeadingHeight = HeadingHeight.Value;
        }

        if (!string.IsNullOrWhiteSpace(FillColor)) {
            options.FillColor = OfficeColor.Parse(FillColor!);
        }

        if (!string.IsNullOrWhiteSpace(LineColor)) {
            options.LineColor = OfficeColor.Parse(LineColor!);
        }

        if (LineWeight.HasValue) {
            options.LineWeight = LineWeight.Value;
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
