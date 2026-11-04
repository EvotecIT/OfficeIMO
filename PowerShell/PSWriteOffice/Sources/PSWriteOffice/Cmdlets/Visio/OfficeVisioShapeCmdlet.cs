using System;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Visio;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Visio;

namespace PSWriteOffice.Cmdlets.Visio;

/// <summary>Shared physical geometry and styling for the four basic Visio shape commands.</summary>
public abstract class OfficeVisioShapeCmdlet : OfficeMutationCmdlet {
    /// <summary>Target page; the current page is used inside a Visio content block.</summary>
    [Parameter(ValueFromPipeline = true)]
    public VisioPage? Page { get; set; }

    /// <summary>Key used by connector and layout commands within the block.</summary>
    [Parameter]
    public string? Key { get; set; }

    /// <summary>Shape pin's X coordinate. A number means points.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength X { get; set; } = OfficeLength.Inches(1);

    /// <summary>Shape pin's Y coordinate. A number means points.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength Y { get; set; } = OfficeLength.Inches(1);

    /// <summary>Physical shape width.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength Width { get; set; } = OfficeLength.Inches(2);

    /// <summary>Physical shape height.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength Height { get; set; } = OfficeLength.Inches(1);

    /// <summary>Optional shape name; otherwise Key is used.</summary>
    [Parameter]
    public string? Name { get; set; }

    /// <summary>Optional universal shape name.</summary>
    [Parameter]
    public string? NameU { get; set; }

    /// <summary>Fill color or semantic role resolved through the default Office theme.</summary>
    [Parameter, OfficeColorTransformation, ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? FillColor { get; set; }

    /// <summary>Line color or semantic role resolved through the default Office theme.</summary>
    [Parameter, OfficeColorTransformation, ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? LineColor { get; set; }

    /// <summary>Physical outline weight.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? LineWeight { get; set; }

    /// <summary>Native line pattern: 0 hides the line, 1 is solid, and 2 through 23 select built-in patterns.</summary>
    [Parameter, ValidateRange(0, 23)]
    public int? LinePattern { get; set; }

    /// <summary>Native fill pattern: 0 has no fill, 1 is solid, and 2 through 40 select built-in patterns.</summary>
    [Parameter, ValidateRange(0, 40)]
    public int? FillPattern { get; set; }

    /// <summary>Rotation in degrees.</summary>
    [Parameter]
    public double? Angle { get; set; }

    /// <summary>Creates the native shape using the bound physical geometry.</summary>
    protected abstract VisioShape CreateShape(VisioPage page);

    /// <inheritdoc />
    protected override void ProcessRecord() {
        VisioDslContext? context = VisioDslContext.Current;
        VisioPage page = Page ?? VisioDslContext.Require(this).RequirePage();
        VisioShape shape = CreateShape(page);
        VisioShapeCommandUtilities.ApplyShapeStyle(shape, Name ?? Key, NameU,
            FillColor?.Resolve(OfficeTheme.Default).ToString(), LineColor?.Resolve(OfficeTheme.Default).ToString(),
            LineWeight?.ToInches(), LinePattern, FillPattern, Angle.HasValue ? Angle.Value * Math.PI / 180D : null);
        context?.RegisterShape(page, Key, shape);
        WritePassThru(shape);
    }
}
