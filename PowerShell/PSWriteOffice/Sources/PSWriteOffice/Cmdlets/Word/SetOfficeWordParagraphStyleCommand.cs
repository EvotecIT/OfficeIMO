using System.Management.Automation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Updates paragraph style, spacing, indentation, and pagination hints.</summary>
/// <example>
///   <summary>Style a paragraph in a Word DSL.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$p = Add-OfficeWordParagraph -Text 'Executive Summary' -PassThru; $p | Set-OfficeWordParagraphStyle -Style Heading1 -KeepWithNext $true</code>
///   <para>Applies a heading style and keeps it with the next paragraph.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeWordParagraphStyle")]
[Alias("WordParagraphStyle")]
[OutputType(typeof(WordParagraph))]
public sealed class SetOfficeWordParagraphStyleCommand : PSCmdlet
{
    /// <summary>Paragraph to update.</summary>
    [Parameter(ValueFromPipeline = true, Position = 0)]
    public WordParagraph? Paragraph { get; set; }

    /// <summary>Paragraph style to apply.</summary>
    [Parameter]
    public WordParagraphStyles? Style { get; set; }

    /// <summary>Paragraph style id to apply.</summary>
    [Parameter]
    public string? StyleId { get; set; }

    /// <summary>Paragraph alignment.</summary>
    [Parameter]
    public WordParagraphAlignment? Alignment { get; set; }

    /// <summary>Vertical character alignment on each line.</summary>
    [Parameter]
    public WordVerticalCharacterAlignment? CharacterAlignment { get; set; }

    /// <summary>Indentation before the paragraph in points.</summary>
    [Parameter]
    public double? IndentationBeforePoints { get; set; }

    /// <summary>Indentation after the paragraph in points.</summary>
    [Parameter]
    public double? IndentationAfterPoints { get; set; }

    /// <summary>First-line indentation in points.</summary>
    [Parameter]
    public double? IndentationFirstLinePoints { get; set; }

    /// <summary>Hanging indentation in points.</summary>
    [Parameter]
    public double? IndentationHangingPoints { get; set; }

    /// <summary>Line spacing in points.</summary>
    [Parameter]
    public double? LineSpacingPoints { get; set; }

    /// <summary>Line spacing before the paragraph in points.</summary>
    [Parameter]
    public double? SpacingBeforePoints { get; set; }

    /// <summary>Line spacing after the paragraph in points.</summary>
    [Parameter]
    public double? SpacingAfterPoints { get; set; }

    /// <summary>Line spacing calculation rule.</summary>
    [Parameter]
    public WordLineSpacingRule? LineSpacingRule { get; set; }

    /// <summary>Start the paragraph on a new page.</summary>
    [Parameter]
    public bool? PageBreakBefore { get; set; }

    /// <summary>Keep this paragraph with the next paragraph.</summary>
    [Parameter]
    public bool? KeepWithNext { get; set; }

    /// <summary>Keep all paragraph lines together.</summary>
    [Parameter]
    public bool? KeepLinesTogether { get; set; }

    /// <summary>Enable widow and orphan control.</summary>
    [Parameter]
    public bool? AvoidWidowAndOrphan { get; set; }

    /// <summary>Paragraph text direction.</summary>
    [Parameter]
    public WordTextDirection? TextDirection { get; set; }

    /// <summary>Set or clear right-to-left paragraph layout.</summary>
    [Parameter]
    public bool? BiDi { get; set; }

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

        if (Style.HasValue) paragraph.Style = Style.Value;
        if (!string.IsNullOrWhiteSpace(StyleId)) paragraph.SetStyleId(StyleId!);
        if (Alignment.HasValue) paragraph.ParagraphAlignment = Alignment.Value;
        if (CharacterAlignment.HasValue) paragraph.VerticalCharacterAlignmentOnLine = CharacterAlignment.Value;
        if (IsBound(nameof(IndentationBeforePoints))) paragraph.IndentationBeforePoints = IndentationBeforePoints;
        if (IsBound(nameof(IndentationAfterPoints))) paragraph.IndentationAfterPoints = IndentationAfterPoints;
        if (IsBound(nameof(IndentationFirstLinePoints))) paragraph.IndentationFirstLinePoints = IndentationFirstLinePoints;
        if (IsBound(nameof(IndentationHangingPoints))) paragraph.IndentationHangingPoints = IndentationHangingPoints;
        if (IsBound(nameof(LineSpacingPoints))) paragraph.LineSpacingPoints = LineSpacingPoints;
        if (IsBound(nameof(SpacingBeforePoints))) paragraph.LineSpacingBeforePoints = SpacingBeforePoints;
        if (IsBound(nameof(SpacingAfterPoints))) paragraph.LineSpacingAfterPoints = SpacingAfterPoints;
        if (LineSpacingRule.HasValue) paragraph.LineSpacingRule = LineSpacingRule.Value;
        if (IsBound(nameof(PageBreakBefore))) paragraph.PageBreakBefore = PageBreakBefore ?? false;
        if (IsBound(nameof(KeepWithNext))) paragraph.KeepWithNext = KeepWithNext ?? false;
        if (IsBound(nameof(KeepLinesTogether))) paragraph.KeepLinesTogether = KeepLinesTogether ?? false;
        if (IsBound(nameof(AvoidWidowAndOrphan))) paragraph.AvoidWidowAndOrphan = AvoidWidowAndOrphan ?? false;
        if (TextDirection.HasValue) paragraph.TextDirection = TextDirection.Value;
        if (IsBound(nameof(BiDi))) paragraph.BiDi = BiDi ?? false;

        if (PassThru.IsPresent)
        {
            WriteObject(paragraph);
        }
    }

    private bool IsBound(string parameterName) => MyInvocation.BoundParameters.ContainsKey(parameterName);
}
