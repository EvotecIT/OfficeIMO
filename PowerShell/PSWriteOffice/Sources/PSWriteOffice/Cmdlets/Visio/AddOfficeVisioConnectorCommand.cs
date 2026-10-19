using System;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Visio;
using PSWriteOffice.Services.Visio;

namespace PSWriteOffice.Cmdlets.Visio;

/// <summary>Adds a connector between two Visio shapes.</summary>
/// <example>
///   <summary>Connect two keyed shapes.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>OfficeVisio -Path .\Flow.vsdx {
///     VisioRectangle -Key source -Text 'Source' -X 1 -Y 4
///     VisioRectangle -Key target -Text 'Target' -X 4 -Y 4
///     VisioConnector -From source -To target -Kind RightAngle -EndArrow Triangle -Label 'sync'
/// }</code>
///   <para>Adds a routed connector between shapes registered in the current DSL scope.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeVisioConnector", DefaultParameterSetName = ByKeyParameterSet)]
[Alias("VisioConnector")]
[OutputType(typeof(VisioConnector))]
public sealed class AddOfficeVisioConnectorCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    private const string ByKeyParameterSet = "ByKey";
    private const string ByShapeParameterSet = "ByShape";

    /// <summary>Target page. Optional inside <c>VisioPage</c> or <c>OfficeVisio</c>.</summary>
    [Parameter(ValueFromPipeline = true)]
    public VisioPage? Page { get; set; }

    /// <summary>Source shape key, id, or name.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ByKeyParameterSet)]
    public string From { get; set; } = string.Empty;

    /// <summary>Target shape key, id, or name.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ByKeyParameterSet)]
    public string To { get; set; } = string.Empty;

    /// <summary>Source shape object.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ByShapeParameterSet)]
    public VisioShape FromShape { get; set; } = null!;

    /// <summary>Target shape object.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ByShapeParameterSet)]
    public VisioShape ToShape { get; set; } = null!;

    /// <summary>Connector kind.</summary>
    [Parameter]
    public ConnectorKind Kind { get; set; } = ConnectorKind.Dynamic;

    /// <summary>Preferred source shape side.</summary>
    [Parameter]
    public VisioSide FromSide { get; set; } = VisioSide.Auto;

    /// <summary>Preferred target shape side.</summary>
    [Parameter]
    public VisioSide ToSide { get; set; } = VisioSide.Auto;

    /// <summary>Connector label.</summary>
    [Parameter]
    public string? Label { get; set; }

    /// <summary>Line color name or hex value.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? LineColor { get; set; }

    /// <summary>Line weight.</summary>
    [Parameter]
    public double? LineWeight { get; set; }

    /// <summary>Native Visio line-pattern index: 0 hides the line, 1 is solid, and 2 through 23 select built-in patterns.</summary>
    [Parameter]
    [ValidateRange(0, 23)]
    public int? LinePattern { get; set; }

    /// <summary>Begin arrow style.</summary>
    [Parameter]
    public EndArrow? BeginArrow { get; set; }

    /// <summary>End arrow style.</summary>
    [Parameter]
    public EndArrow? EndArrow { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var context = VisioDslContext.Current;
        var page = Page ?? (context ?? VisioDslContext.Require(this)).RequirePage();
        var fromShape = ParameterSetName == ByShapeParameterSet
            ? FromShape
            : ResolveShape(context, page, From);
        var toShape = ParameterSetName == ByShapeParameterSet
            ? ToShape
            : ResolveShape(context, page, To);

        var connector = page.AddConnector(fromShape, toShape, Kind, FromSide, ToSide);
        VisioShapeCommandUtilities.ApplyConnectorStyle(connector, LineColor, LineWeight, LinePattern, BeginArrow, EndArrow, Label);
        WritePassThru(connector);
    }

    private VisioShape ResolveShape(VisioDslContext? context, VisioPage page, string value) {
        if (context != null) {
            return context.ResolveShape(page, value);
        }

        var shape = page.AllShapes().FirstOrDefault(candidate =>
            string.Equals(candidate.Id, value, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(candidate.Name, value, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(candidate.NameU, value, StringComparison.OrdinalIgnoreCase));

        return shape ?? throw new PSInvalidOperationException($"Visio shape '{value}' was not found on the target page.");
    }
}
