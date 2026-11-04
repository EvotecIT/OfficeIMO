using System;
using System.Linq;
using System.Management.Automation;
using OfficeIMO;
using PSWriteOffice.Foundation;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Stencils;
using PSWriteOffice.Services.Visio;

namespace PSWriteOffice.Cmdlets.Visio;

/// <summary>Adds a stencil shape to the current Visio page.</summary>
/// <example>
///   <summary>Add built-in flowchart stencil shapes.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>OfficeVisio -Path .\StencilFlow.vsdx -UseMastersByDefault {
///     VisioStencilImport -BuiltIn Flowchart -Name Flow -Default
///     VisioStencil -Catalog Flow -Stencil process -Key intake -Text 'Intake' -X 1.5 -Y 4
/// }</code>
///   <para>Registers a built-in catalog and places a stencil shape on the active page.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeVisioStencilShape", DefaultParameterSetName = "Catalog")]
[OutputType(typeof(VisioShape))]
[Alias("VisioStencil") ]
public sealed class AddOfficeVisioStencilShapeCommand : OfficeVisioTextCmdlet {
    private const string BuiltInParameterSet = "BuiltIn";

    /// <summary>Target page. Optional inside <c>VisioPage</c> or <c>OfficeVisio</c>.</summary>
    [Parameter(ValueFromPipeline = true)]
    public VisioPage? Page { get; set; }

    /// <summary>A stencil catalog object or a name registered in the current Visio block.</summary>
    [Parameter(ParameterSetName = "Catalog")]
    public object? Catalog { get; set; }

    /// <summary>Built-in OfficeIMO stencil catalog containing the shape.</summary>
    [Parameter(Mandatory = true, ParameterSetName = BuiltInParameterSet)]
    public OfficeVisioBuiltInStencilCatalog BuiltIn { get; set; } = OfficeVisioBuiltInStencilCatalog.All;

    /// <summary>Stencil id, name, master name, keyword, alias, or tag.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Stencil { get; set; } = string.Empty;

    /// <summary>DSL key used by connector commands.</summary>
    [Parameter]
    public string? Key { get; set; }

    /// <summary>X coordinate of the stencil shape center.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength X { get; set; } = OfficeLength.Inches(1);

    /// <summary>Y coordinate of the stencil shape center.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength Y { get; set; } = OfficeLength.Inches(1);

    /// <summary>Optional shape width. Omit to use the stencil default width.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? Width { get; set; }

    /// <summary>Optional shape height. Omit to use the stencil default height.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? Height { get; set; }

    /// <summary>Text placed inside the shape. Omit to use the stencil display name.</summary>
    [Parameter(Position = 1)]
    public string? Text { get; set; }

    /// <summary>Optional shape name.</summary>
    [Parameter]
    public string? Name { get; set; }

    /// <summary>Optional universal shape name.</summary>
    [Parameter]
    public string? NameU { get; set; }

    /// <summary>Fill color name or hex value.</summary>
    [Parameter]
    [OfficeColorTransformation, ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? FillColor { get; set; }

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

    /// <summary>Native Visio fill-pattern index: 0 has no fill, 1 is solid, and 2 through 40 select built-in patterns.</summary>
    [Parameter]
    [ValidateRange(0, 40)]
    public int? FillPattern { get; set; }

    /// <summary>Shape angle in degrees.</summary>
    [Parameter]
    public double? Angle { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var context = VisioDslContext.Current;
        var page = Page ?? VisioDslContext.Require(this).RequirePage();
        var catalog = ResolveCatalog(context);
        var stencil = catalog.Get(Stencil);
        var id = string.IsNullOrWhiteSpace(Key) ? CreateUniqueShapeId(page, stencil.Id) : Key!;
        var x = X;
        var y = Y;
        var shape = Width.HasValue || Height.HasValue
            ? page.AddStencilShape(
                catalog,
                Stencil,
                id,
                x,
                y,
                Width.HasValue ? Width.Value : ToLength(stencil.DefaultWidth, stencil.DefaultUnit ?? page.DefaultUnit),
                Height.HasValue ? Height.Value : ToLength(stencil.DefaultHeight, stencil.DefaultUnit ?? page.DefaultUnit),
                Text)
            : page.AddStencilShape(catalog, Stencil, id, x, y, Text);

        VisioShapeCommandUtilities.ApplyShapeStyle(shape, Name ?? Key, NameU, FillColor?.Resolve(OfficeTheme.Default).ToString(), LineColor?.Resolve(OfficeTheme.Default).ToString(), LineWeight?.ToInches(), LinePattern, FillPattern, Angle.HasValue ? Angle.Value * Math.PI / 180D : null);
        ApplyTextStyle(shape);
        context?.RegisterShape(page, Key, shape);
        WritePassThru(shape);
    }

    private VisioStencilCatalog ResolveCatalog(VisioDslContext? context) {
        object? value = Catalog is PSObject wrapped ? wrapped.BaseObject : Catalog;
        if (value is VisioStencilCatalog supplied) return supplied;
        if (value != null && value is not string) throw new PSArgumentException("Catalog accepts a VisioStencilCatalog or a registered name.");
        string? name = value as string;

        if (ParameterSetName == BuiltInParameterSet) {
            return VisioStencilCommandUtilities.GetBuiltInCatalog(BuiltIn);
        }

        if (context != null) {
            return context.ResolveStencilCatalog(name);
        }

        return string.IsNullOrWhiteSpace(name)
            ? VisioStencils.All
            : throw new PSInvalidOperationException("A named stencil catalog can only be resolved inside New-OfficeVisio.");
    }

    private static string CreateUniqueShapeId(VisioPage page, string baseId) {
        string stem = string.IsNullOrWhiteSpace(baseId) ? "stencil" : baseId.Trim();
        var existingIds = page.AllShapes()
            .Select(shape => shape.Id)
            .Where(id => !string.IsNullOrWhiteSpace(id))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        if (!existingIds.Contains(stem)) {
            return stem;
        }

        for (int index = 2; ; index++) {
            string candidate = stem + "-" + index.ToString(System.Globalization.CultureInfo.InvariantCulture);
            if (!existingIds.Contains(candidate)) {
                return candidate;
            }
        }
    }

    private static OfficeIMO.OfficeLength ToLength(double value, VisioMeasurementUnit unit) => unit switch {
        VisioMeasurementUnit.Centimeters => OfficeIMO.OfficeLength.Centimeters(value),
        VisioMeasurementUnit.Millimeters => OfficeIMO.OfficeLength.Millimeters(value),
        _ => OfficeIMO.OfficeLength.Inches(value)
    };
}
