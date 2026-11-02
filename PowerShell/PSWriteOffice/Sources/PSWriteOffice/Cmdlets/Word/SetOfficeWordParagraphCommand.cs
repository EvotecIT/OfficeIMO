using System;
using System.Management.Automation;
using OfficeIMO;
using PSWriteOffice.Foundation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Updates paragraph style, spacing, indentation, and pagination hints.</summary>
/// <example>
///   <summary>Style a paragraph in a Word DSL.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$p = Add-OfficeWordParagraph -Text 'Executive Summary' -PassThru; $p | Set-OfficeWordParagraph -Style Heading1 -KeepWithNext</code>
///   <para>Applies a heading style and keeps it with the next paragraph.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeWordParagraph", DefaultParameterSetName = "Default")]
[OutputType(typeof(WordParagraph))]
public sealed class SetOfficeWordParagraphCommand : OfficeCmdlet
{
    /// <summary>Paragraph to update.</summary>
    [Parameter(ValueFromPipeline = true, Position = 0, ParameterSetName = "Default")]
    public WordParagraph? Paragraph { get; set; }

    /// <summary>Paragraph style to apply.</summary>
    [Parameter]
    public WordParagraphStyles? Style { get; set; }

    /// <summary>Paragraph style id to apply.</summary>
    [Parameter]
    public string? StyleId { get; set; }

    /// <summary>Paragraph alignment.</summary>
    [Parameter]
    public OfficeAlignment? Alignment { get; set; }

    /// <summary>Vertical character alignment on each line.</summary>
    [Parameter]
    public WordVerticalCharacterAlignment? CharacterAlignment { get; set; }

    /// <summary>Indentation before the paragraph; a bare number means points.</summary>
    [Parameter, OfficeLengthTransformation(AllowNull = true)]
    public OfficeLength? IndentationBefore { get; set; }

    /// <summary>Indentation after the paragraph; a bare number means points.</summary>
    [Parameter, OfficeLengthTransformation(AllowNull = true)]
    public OfficeLength? IndentationAfter { get; set; }

    /// <summary>First-line indentation; a bare number means points.</summary>
    [Parameter, OfficeLengthTransformation(AllowNull = true)]
    public OfficeLength? IndentationFirstLine { get; set; }

    /// <summary>Hanging indentation; a bare number means points.</summary>
    [Parameter, OfficeLengthTransformation(AllowNull = true)]
    public OfficeLength? IndentationHanging { get; set; }

    /// <summary>Physical line height; bare numbers mean points. Null clears the direct line-spacing override.</summary>
    [Parameter, OfficeLengthTransformation(AllowNull = true)]
    public OfficeLength? LineSpacing { get; set; }

    /// <summary>Line spacing before the paragraph; a bare number means points.</summary>
    [Parameter, OfficeLengthTransformation(AllowNull = true)]
    public OfficeLength? SpacingBefore { get; set; }

    /// <summary>Line spacing after the paragraph; a bare number means points.</summary>
    [Parameter, OfficeLengthTransformation(AllowNull = true)]
    public OfficeLength? SpacingAfter { get; set; }

    /// <summary>Line spacing calculation rule.</summary>
    [Parameter]
    public WordLineSpacingRule? LineSpacingRule { get; set; }

    /// <summary>Start the paragraph on a new page.</summary>
    [Parameter]
    public SwitchParameter PageBreakBefore { get; set; }

    /// <summary>Keep this paragraph with the next paragraph.</summary>
    [Parameter]
    public SwitchParameter KeepWithNext { get; set; }

    /// <summary>Keep all paragraph lines together.</summary>
    [Parameter]
    public SwitchParameter KeepLinesTogether { get; set; }

    /// <summary>Enable widow and orphan control.</summary>
    [Parameter]
    public SwitchParameter AvoidWidowAndOrphan { get; set; }

    /// <summary>Paragraph text direction.</summary>
    [Parameter]
    public WordTextDirection? TextDirection { get; set; }

    /// <summary>Set or clear right-to-left paragraph layout.</summary>
    [Parameter]
    public SwitchParameter BiDi { get; set; }

    /// <summary>Emit the updated paragraph.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var paragraph = Paragraph ?? WordDslContext.Current?.CurrentParagraph;
        if (paragraph == null)
        {
            throw new PSInvalidOperationException("Provide a Word paragraph or run this command inside a Word paragraph DSL scope.");
        }

        WordParagraphAlignment? alignment = Alignment switch {
            null => null,
            OfficeAlignment.Left => WordParagraphAlignment.Left,
            OfficeAlignment.Start => WordParagraphAlignment.Start,
            OfficeAlignment.Right => WordParagraphAlignment.Right,
            OfficeAlignment.End => WordParagraphAlignment.End,
            OfficeAlignment.Center => WordParagraphAlignment.Center,
            OfficeAlignment.Justify => WordParagraphAlignment.Both,
            _ => throw new PSArgumentException("Alignment must be a horizontal paragraph alignment.")
        };
        bool spacingBound = IsBound(nameof(LineSpacing)) || LineSpacingRule.HasValue;
        WordLineSpacing? spacing = paragraph.LineSpacing;
        if (IsBound(nameof(LineSpacing))) {
            if (!LineSpacing.HasValue) {
                if (LineSpacingRule.HasValue) throw new PSArgumentException("Provide a line height with LineSpacingRule.");
                spacing = null;
            } else {
                WordLineSpacingRule rule = LineSpacingRule ?? (spacing?.Rule == WordLineSpacingRule.AtLeast ? WordLineSpacingRule.AtLeast : WordLineSpacingRule.Exact);
                spacing = rule switch {
                    WordLineSpacingRule.Exact => WordLineSpacing.Exact(LineSpacing.Value),
                    WordLineSpacingRule.AtLeast => WordLineSpacing.AtLeast(LineSpacing.Value),
                    _ => throw new PSArgumentException("LineSpacing is a physical height; select Exact or AtLeast instead of Auto.")
                };
            }
        } else if (LineSpacingRule.HasValue) {
            spacing = LineSpacingRule.Value switch {
                WordLineSpacingRule.Auto when spacing?.MultipleValue is double multiple => WordLineSpacing.Multiple(multiple),
                WordLineSpacingRule.Exact when spacing?.Height is OfficeLength height => WordLineSpacing.Exact(height),
                WordLineSpacingRule.AtLeast when spacing?.Height is OfficeLength minimum => WordLineSpacing.AtLeast(minimum),
                _ => throw new PSArgumentException("Provide LineSpacing when changing between a line multiple and a physical height.")
            };
        }

        if (Style.HasValue) paragraph.Style = Style.Value;
        if (!string.IsNullOrWhiteSpace(StyleId)) paragraph.SetStyleId(StyleId!);
        if (Alignment.HasValue) paragraph.ParagraphAlignment = alignment;
        if (CharacterAlignment.HasValue) paragraph.VerticalCharacterAlignmentOnLine = CharacterAlignment.Value;
        if (IsBound(nameof(IndentationBefore)) && IndentationBefore.HasValue) paragraph.IndentationBefore = IndentationBefore.Value;
        if (IsBound(nameof(IndentationAfter)) && IndentationAfter.HasValue) paragraph.IndentationAfter = IndentationAfter.Value;
        if (IsBound(nameof(IndentationFirstLine)) && IndentationFirstLine.HasValue) paragraph.IndentationFirstLine = IndentationFirstLine.Value;
        if (IsBound(nameof(IndentationHanging)) && IndentationHanging.HasValue) paragraph.IndentationHanging = IndentationHanging.Value;
        if (spacingBound) paragraph.LineSpacing = spacing;
        if (IsBound(nameof(SpacingBefore)) && SpacingBefore.HasValue) paragraph.LineSpacingBefore = SpacingBefore.Value;
        if (IsBound(nameof(SpacingAfter)) && SpacingAfter.HasValue) paragraph.LineSpacingAfter = SpacingAfter.Value;
        if (BoundSwitch(nameof(PageBreakBefore)) is bool pagebreakbefore) paragraph.PageBreakBefore = pagebreakbefore;
        if (BoundSwitch(nameof(KeepWithNext)) is bool keepwithnext) paragraph.KeepWithNext = keepwithnext;
        if (BoundSwitch(nameof(KeepLinesTogether)) is bool keeplinestogether) paragraph.KeepLinesTogether = keeplinestogether;
        if (BoundSwitch(nameof(AvoidWidowAndOrphan)) is bool avoidwidowandorphan) paragraph.AvoidWidowAndOrphan = avoidwidowandorphan;
        if (TextDirection.HasValue) paragraph.TextDirection = TextDirection.Value;
        if (BoundSwitch(nameof(BiDi)) is bool bidi) paragraph.BiDi = bidi;

        if (PassThru.IsPresent)
        {
            WriteObject(paragraph);
        }
    }

    private bool IsBound(string parameterName) => MyInvocation.BoundParameters.ContainsKey(parameterName);
}
