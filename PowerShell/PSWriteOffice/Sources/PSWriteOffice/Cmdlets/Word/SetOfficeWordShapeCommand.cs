using OfficeIMO;
using PSWriteOffice.Foundation;
using System.Management.Automation;
using OfficeIMO.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Updates OfficeIMO Word shape metadata, sizing, and colors.</summary>
/// <example>
///   <summary>Restyle callout shapes in an opened report.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$doc = Open-OfficeWord -Path .\Report.docx
/// $doc |
///     Get-OfficeWordShape |
///     Set-OfficeWordShape -FillColor '#fff7e6' -StrokeColor '#fa8c16' -StrokeWidth 1.25 -Description 'Highlighted callout'
/// $doc | Save-OfficeWord -Path .\Report-Shapes.docx</code>
///   <para>Updates OfficeIMO shape objects through the pipeline and persists the document with the standard save command.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeWordShape", DefaultParameterSetName = "Default")]
[OutputType(typeof(WordShape))]
public sealed class SetOfficeWordShapeCommand : PSCmdlet
{
    /// <summary>Shape to update.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0, ParameterSetName = "Default")]
    public WordShape Shape { get; set; } = null!;

    /// <summary>Shape width in points.</summary>
    [Parameter, OfficeLengthTransformation] public OfficeLength? Width { get; set; }

    /// <summary>Shape height in points.</summary>
    [Parameter, OfficeLengthTransformation] public OfficeLength? Height { get; set; }

    /// <summary>Anchored left position in points.</summary>
    [Parameter, AllowNull, OfficeLengthTransformation(AllowNull = true)] public OfficeLength? X { get; set; }

    /// <summary>Anchored top position in points.</summary>
    [Parameter, AllowNull, OfficeLengthTransformation(AllowNull = true)] public OfficeLength? Y { get; set; }

    /// <summary>Shape rotation in degrees.</summary>
    [Parameter] public double? Rotation { get; set; }

    /// <summary>Fill color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? FillColor { get; set; }

    /// <summary>Stroke color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? StrokeColor { get; set; }

    /// <summary>Stroke width in points.</summary>
    [Parameter, AllowNull, OfficeLengthTransformation(AllowNull = true)] public OfficeLength? StrokeWidth { get; set; }

    /// <summary>Whether the shape stroke is enabled.</summary>
    [Parameter] public SwitchParameter Stroked { get; set; }

    /// <summary>Shape z-index.</summary>
    [Parameter] public int? ZIndex { get; set; }

    /// <summary>Shape title metadata.</summary>
    [Parameter] public string? Title { get; set; }

    /// <summary>Shape alternate text metadata.</summary>
    [Parameter] public string? Description { get; set; }

    /// <summary>Whether the shape is hidden.</summary>
    [Parameter] public SwitchParameter Hidden { get; set; }

    /// <summary>Emit the updated shape.</summary>
    [Parameter] public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (Shape == null)
        {
            return;
        }

        if (MyInvocation.BoundParameters.ContainsKey(nameof(Width)) && Width.HasValue) Shape.Width = Width.Value;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(Height)) && Height.HasValue) Shape.Height = Height.Value;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(X))) Shape.Left = X;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(Y))) Shape.Top = Y;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(Rotation))) Shape.Rotation = Rotation;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(FillColor))) Shape.FillColor = OfficeColor.ParseHex(WordRunColor.Create(FillColor!.Value, Shape._document).Val!.Value!);
        if (MyInvocation.BoundParameters.ContainsKey(nameof(StrokeColor))) Shape.StrokeColor = OfficeColor.ParseHex(WordRunColor.Create(StrokeColor!.Value, Shape._document).Val!.Value!);
        if (MyInvocation.BoundParameters.ContainsKey(nameof(StrokeWidth))) Shape.StrokeWeight = StrokeWidth;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(Stroked))) Shape.Stroked = Stroked.IsPresent;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(ZIndex))) Shape.ZIndex = ZIndex;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(Title))) Shape.Title = Title;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(Description))) Shape.Description = Description;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(Hidden))) Shape.Hidden = Hidden.IsPresent;

        if (PassThru.IsPresent)
        {
            WriteObject(Shape);
        }
    }
}
