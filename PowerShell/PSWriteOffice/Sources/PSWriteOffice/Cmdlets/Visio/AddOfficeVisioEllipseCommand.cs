using System.Management.Automation;
using OfficeIMO.Visio;
using PSWriteOffice.Services.Visio;

namespace PSWriteOffice.Cmdlets.Visio;

/// <summary>Adds an ellipse shape to the current Visio page.</summary>
/// <example>
///   <summary>Add a start node.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>OfficeVisio -Path .\Flow.vsdx {
///     VisioEllipse -Key start -Text 'Start' -X 1 -Y 5 -Width 1.4 -Height 0.7 -FillColor '#DCFCE7'
/// }</code>
///   <para>Adds an ellipse to the active Visio page.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeVisioEllipse")]
[Alias("VisioEllipse")]
[OutputType(typeof(VisioShape))]
public sealed class AddOfficeVisioEllipseCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    /// <summary>Target page. Optional inside <c>VisioPage</c> or <c>OfficeVisio</c>.</summary>
    [Parameter(ValueFromPipeline = true)]
    public VisioPage? Page { get; set; }

    /// <summary>DSL key used by connector commands.</summary>
    [Parameter]
    public string? Key { get; set; }

    /// <summary>X coordinate of the shape origin.</summary>
    [Parameter]
    public double X { get; set; } = 1;

    /// <summary>Y coordinate of the shape origin.</summary>
    [Parameter]
    public double Y { get; set; } = 1;

    /// <summary>Shape width.</summary>
    [Parameter]
    public double Width { get; set; } = 2;

    /// <summary>Shape height.</summary>
    [Parameter]
    public double Height { get; set; } = 1;

    /// <summary>Text placed inside the shape.</summary>
    [Parameter(Position = 0)]
    public string? Text { get; set; }

    /// <summary>Measurement unit for coordinates and dimensions.</summary>
    [Parameter]
    public VisioMeasurementUnit Unit { get; set; } = VisioMeasurementUnit.Inches;

    /// <summary>Optional shape name.</summary>
    [Parameter]
    public string? Name { get; set; }

    /// <summary>Fill color name or hex value.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? FillColor { get; set; }

    /// <summary>Line color name or hex value.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? LineColor { get; set; }

    /// <summary>Line weight.</summary>
    [Parameter]
    public double? LineWeight { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var context = VisioDslContext.Current;
        var page = Page ?? VisioDslContext.Require(this).RequirePage();
        var shape = page.AddEllipse(X, Y, Width, Height, Text, Unit);
        VisioShapeCommandUtilities.ApplyShapeStyle(shape, Name ?? Key, null, FillColor, LineColor, LineWeight, null, null, null);
        context?.RegisterShape(page, Key, shape);
        WritePassThru(shape);
    }
}
