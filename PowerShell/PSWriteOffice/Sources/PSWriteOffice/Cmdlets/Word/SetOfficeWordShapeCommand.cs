using System.Management.Automation;
using OfficeIMO.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Updates OfficeIMO Word shape metadata, sizing, and colors.</summary>
/// <example>
///   <summary>Restyle callout shapes in an opened report.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$doc = Get-OfficeWord -Path .\Report.docx
/// $doc |
///     Get-OfficeWordShape |
///     Set-OfficeWordShape -FillColor '#fff7e6' -StrokeColor '#fa8c16' -StrokeWidth 1.25 -Description 'Highlighted callout'
/// $doc | Save-OfficeWord -Path .\Report-Shapes.docx</code>
///   <para>Updates OfficeIMO shape objects through the pipeline and persists the document with the standard save command.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeWordShape")]
[Alias("WordShapeStyle")]
[OutputType(typeof(WordShape))]
public sealed class SetOfficeWordShapeCommand : PSCmdlet
{
    /// <summary>Shape to update.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
    public WordShape Shape { get; set; } = null!;

    /// <summary>Shape width in points.</summary>
    [Parameter] public double? Width { get; set; }

    /// <summary>Shape height in points.</summary>
    [Parameter] public double? Height { get; set; }

    /// <summary>Anchored left position in points.</summary>
    [Parameter] public double? Left { get; set; }

    /// <summary>Anchored top position in points.</summary>
    [Parameter] public double? Top { get; set; }

    /// <summary>Shape rotation in degrees.</summary>
    [Parameter] public double? Rotation { get; set; }

    /// <summary>Fill color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? FillColor { get; set; }

    /// <summary>Stroke color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? StrokeColor { get; set; }

    /// <summary>Stroke width in points.</summary>
    [Parameter] public double? StrokeWidth { get; set; }

    /// <summary>Whether the shape stroke is enabled.</summary>
    [Parameter] public bool? Stroked { get; set; }

    /// <summary>Shape z-index.</summary>
    [Parameter] public int? ZIndex { get; set; }

    /// <summary>Shape title metadata.</summary>
    [Parameter] public string? Title { get; set; }

    /// <summary>Shape alternate text metadata.</summary>
    [Parameter] public string? Description { get; set; }

    /// <summary>Whether the shape is hidden.</summary>
    [Parameter] public bool? Hidden { get; set; }

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
        if (MyInvocation.BoundParameters.ContainsKey(nameof(Left))) Shape.Left = Left;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(Top))) Shape.Top = Top;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(Rotation))) Shape.Rotation = Rotation;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(FillColor))) Shape.FillColorHex = FillColor ?? string.Empty;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(StrokeColor))) Shape.StrokeColorHex = StrokeColor ?? string.Empty;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(StrokeWidth))) Shape.StrokeWeight = StrokeWidth;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(Stroked))) Shape.Stroked = Stroked;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(ZIndex))) Shape.ZIndex = ZIndex;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(Title))) Shape.Title = Title;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(Description))) Shape.Description = Description;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(Hidden))) Shape.Hidden = Hidden;

        if (PassThru.IsPresent)
        {
            WriteObject(Shape);
        }
    }
}
