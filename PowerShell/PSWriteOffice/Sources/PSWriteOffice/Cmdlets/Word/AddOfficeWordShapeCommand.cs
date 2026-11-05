using OfficeIMO;
using PSWriteOffice.Foundation;
using System;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Adds a basic OfficeIMO Word shape to the current paragraph.</summary>
/// <example>
///   <summary>Add a callout shape to a report section.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeWord -Path .\StatusReport.docx {
///     Add-OfficeWordParagraph -Text 'Release readiness'
///     Add-OfficeWordShape -Type Rectangle -Width 220 -Height 56 -FillColor '#e6fffb' -StrokeColor '#08979c' -StrokeWidth 1.5 -Title 'Status callout' -Description 'Release readiness callout'
/// }</code>
///   <para>Creates an OfficeIMO Word shape in the current paragraph and sets basic visual and accessibility metadata.</para>
/// </example>
/// <example>
///   <summary>Add an anchored background shape.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeWord -Path .\Appendix.docx {
///     Add-OfficeWordParagraph -Text 'Appendix A'
///     Add-OfficeWordShape -Type Rectangle -Width 480 -Height 36 -X 36 -Y 72 -FillColor '#f0f5ff' -StrokeColor '#adc6ff'
/// }</code>
///   <para>Positions a shape with explicit offsets when the OfficeIMO anchored-shape API is desired.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeWordShape", DefaultParameterSetName = "Default")]
[OutputType(typeof(WordShape))]
[Alias("WordShape") ]
public sealed class AddOfficeWordShapeCommand : PSCmdlet
{
    /// <summary>Shape type to add.</summary>
    [Parameter(ParameterSetName = "Default")]
    public WordShapeType Type { get; set; } = WordShapeType.Rectangle;

    /// <summary>Width in points.</summary>
    [Parameter]
    [OfficeLengthTransformation]
    public OfficeLength Width { get; set; } = OfficeLength.Points(120);

    /// <summary>Height in points.</summary>
    [Parameter]
    [OfficeLengthTransformation]
    public OfficeLength Height { get; set; } = OfficeLength.Points(48);

    /// <summary>Anchored left position in points. When omitted, the shape is inline.</summary>
    [Parameter]
    [OfficeLengthTransformation]
    public OfficeLength? X { get; set; }

    /// <summary>Anchored top position in points. When omitted, the shape is inline.</summary>
    [Parameter]
    [OfficeLengthTransformation]
    public OfficeLength? Y { get; set; }

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
    [Parameter]
    [OfficeLengthTransformation]
    public OfficeLength? StrokeWidth { get; set; }

    /// <summary>Optional title metadata.</summary>
    [Parameter]
    public string? Title { get; set; }

    /// <summary>Optional alternate text metadata.</summary>
    [Parameter]
    public string? Description { get; set; }

    /// <summary>Emit the created shape.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <summary>Native paragraph; omission uses the active Word block.</summary>
    [Parameter(ValueFromPipeline = true)]
    public WordParagraph? Paragraph { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var paragraph = Paragraph ?? PSWriteOffice.Foundation.OfficeDocumentContext.Find<WordParagraph>();
        if (paragraph == null) { var context = WordDslContext.Require(this); paragraph = context.CurrentParagraph ?? context.RequireParagraphHost().AddParagraph(); }

        var shape = X.HasValue || Y.HasValue
            ? WordShape.AddDrawingShapeAnchored(paragraph, Type, Width, Height, X ?? OfficeLength.Zero, Y ?? OfficeLength.Zero)
            : WordShape.AddDrawingShape(paragraph, Type, Width, Height);

        ApplyShape(shape);

        if (PassThru.IsPresent)
        {
            WriteObject(shape);
        }
    }

    private void ApplyShape(WordShape shape)
    {
        if (FillColor.HasValue)
        {
            shape.FillColor = FillColor.Value;
        }

        if (StrokeColor.HasValue)
        {
            shape.StrokeColor = StrokeColor.Value;
        }

        if (StrokeWidth.HasValue)
        {
            shape.StrokeWeight = StrokeWidth.Value;
        }

        if (Title != null)
        {
            shape.Title = Title;
        }

        if (Description != null)
        {
            shape.Description = Description;
        }
    }
}
