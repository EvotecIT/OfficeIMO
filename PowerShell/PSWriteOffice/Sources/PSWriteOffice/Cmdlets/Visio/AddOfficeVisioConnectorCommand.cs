using System;
using System.Linq;
using System.Management.Automation;
using OfficeIMO;
using PSWriteOffice.Foundation;
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
[Cmdlet(VerbsCommon.Add, "OfficeVisioConnector", DefaultParameterSetName = "Page")]
[OutputType(typeof(VisioConnector))]
[Alias("VisioConnector") ]
public sealed class AddOfficeVisioConnectorCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {

    /// <summary>Target page. Optional inside <c>VisioPage</c> or <c>OfficeVisio</c>.</summary>
    [Parameter(ValueFromPipeline = true)]
    public VisioPage? Page { get; set; }

    /// <summary>Source shape object or a registered key, id or name.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "Page")]
    public object From { get; set; } = null!;
    /// <summary>Target shape object or a registered key, id or name.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "Page")]
    public object To { get; set; } = null!;

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
    [OfficeColorTransformation, ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? LineColor { get; set; }

    /// <summary>Line weight.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? LineWeight { get; set; }

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
        var fromShape = ResolveShapeInput(context, page, From);
        var toShape = ResolveShapeInput(context, page, To);

        var connector = page.AddConnector(fromShape, toShape, Kind, FromSide, ToSide);
        VisioShapeCommandUtilities.ApplyConnectorStyle(connector, LineColor?.Resolve(OfficeTheme.Default).ToString(), LineWeight?.ToInches(), LinePattern, BeginArrow, EndArrow, Label);
        WritePassThru(connector);
    }

    private VisioShape ResolveShapeInput(VisioDslContext? context, VisioPage page, object input) {
        object value = input is PSObject wrapped ? wrapped.BaseObject : input;
        if (value is VisioShape shape) {
            if (!page.AllShapes().Contains(shape)) throw new PSArgumentException("The connector shape must belong to its target page.");
            return shape;
        }
        if (value is string key) return ResolveShape(context, page, key);
        throw new PSArgumentException("From and To accept a VisioShape or a shape key, id or name.");
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
