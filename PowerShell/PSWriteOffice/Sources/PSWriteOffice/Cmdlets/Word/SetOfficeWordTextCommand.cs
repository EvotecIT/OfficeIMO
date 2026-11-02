using System;
using System.Management.Automation;
using OfficeIMO;
using PSWriteOffice.Foundation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Updates styling on Word text.</summary>
/// <example>
///   <summary>Style text inside a Word paragraph.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>WordParagraph -Text 'Warning' -PassThru | Set-OfficeWordText -Bold -Color '#C00000'</code>
///   <para>Applies bold red styling to the supplied paragraph or scoped run.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeWordText", DefaultParameterSetName = "Default")]
[OutputType(typeof(WordParagraph))]
public sealed class SetOfficeWordTextCommand : OfficeCmdlet
{
    /// <summary>Word text item to update.</summary>
    [Parameter(ValueFromPipeline = true, Position = 0, ParameterSetName = "Default")]
    public WordParagraph? Paragraph { get; set; }

    /// <summary>Replace the text.</summary>
    [Parameter]
    public string? Text { get; set; }

    /// <summary>Character style to apply.</summary>
    [Parameter]
    public WordCharacterStyles? Style { get; set; }

    /// <summary>Character style id to apply.</summary>
    [Parameter]
    public string? StyleId { get; set; }

    /// <summary>Set or clear bold formatting.</summary>
    [Parameter]
    public SwitchParameter Bold { get; set; }

    /// <summary>Set or clear italic formatting.</summary>
    [Parameter]
    public SwitchParameter Italic { get; set; }

    /// <summary>Underline style to apply.</summary>
    [Parameter]
    public WordUnderlineStyle? Underline { get; set; }

    /// <summary>Text color. Named colors and hexadecimal values are accepted; null clears the direct override.</summary>
    [Parameter]
    [OfficeColorTransformation(AllowNull = true)]
    public OfficeColor? Color { get; set; }

    /// <summary>Font size; bare numbers mean points. Null clears the direct size override.</summary>
    [Parameter, OfficeLengthTransformation(AllowNull = true)]
    public OfficeLength? Size { get; set; }

    /// <summary>Font family name.</summary>
    [Parameter]
    public string? Font { get; set; }

    /// <summary>Highlight color.</summary>
    [Parameter]
    public WordHighlightColor? Highlight { get; set; }

    /// <summary>Set or clear strikethrough.</summary>
    [Parameter]
    public SwitchParameter Strikethrough { get; set; }

    /// <summary>Set or clear double strikethrough.</summary>
    [Parameter]
    public SwitchParameter DoubleStrikethrough { get; set; }

    /// <summary>Capitalization style.</summary>
    [Parameter]
    public WordCapsStyle? CapsStyle { get; set; }

    /// <summary>Character spacing in twentieths of a point.</summary>
    [Parameter]
    public int? Spacing { get; set; }

    /// <summary>Set or clear outline effect.</summary>
    [Parameter]
    public SwitchParameter Outline { get; set; }

    /// <summary>Set or clear shadow effect.</summary>
    [Parameter]
    public SwitchParameter Shadow { get; set; }

    /// <summary>Set or clear emboss effect.</summary>
    [Parameter]
    public SwitchParameter Emboss { get; set; }

    /// <summary>Emit the updated text item.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var text = Paragraph ?? WordDslContext.Current?.CurrentParagraph;
        if (text == null)
        {
            throw new PSInvalidOperationException("Provide a Word text item or run this command inside a Word paragraph DSL scope.");
        }

        if (IsBound(nameof(Text))) text.Text = Text ?? string.Empty;
        if (Style.HasValue) text.SetCharacterStyle(Style.Value);
        if (!string.IsNullOrWhiteSpace(StyleId)) text.SetCharacterStyleId(StyleId!);
        if (BoundSwitch(nameof(Bold)) is bool bold) text.Bold = bold;
        if (BoundSwitch(nameof(Italic)) is bool italic) text.Italic = italic;
        if (Underline.HasValue) text.Underline = Underline.Value;
        if (IsBound(nameof(Color))) { if (Color.HasValue) text.Color = Color; else text.ColorHex = string.Empty; }
        if (IsBound(nameof(Size))) text.FontSize = Size;
        if (IsBound(nameof(Font))) text.FontFamily = Font;
        if (Highlight.HasValue) text.Highlight = Highlight.Value;
        if (BoundSwitch(nameof(Strikethrough)) is bool strikethrough) text.Strike = strikethrough;
        if (BoundSwitch(nameof(DoubleStrikethrough)) is bool doublestrikethrough) text.DoubleStrike = doublestrikethrough;
        if (CapsStyle.HasValue) text.CapsStyle = CapsStyle.Value;
        if (IsBound(nameof(Spacing))) text.Spacing = Spacing.HasValue ? OfficeIMO.OfficeLength.Emus((long)Spacing.Value * 635) : (OfficeIMO.OfficeLength?)null;
        if (BoundSwitch(nameof(Outline)) is bool outline) text.Outline = outline;
        if (BoundSwitch(nameof(Shadow)) is bool shadow) text.Shadow = shadow;
        if (BoundSwitch(nameof(Emboss)) is bool emboss) text.Emboss = emboss;

        if (PassThru.IsPresent)
        {
            WriteObject(text);
        }
    }

    private bool IsBound(string parameterName) => MyInvocation.BoundParameters.ContainsKey(parameterName);
}
