using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Updates styling on Word text.</summary>
/// <example>
///   <summary>Style text inside a Word paragraph.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>WordParagraph -Text 'Warning' -PassThru | Set-OfficeWordTextStyle -Bold $true -Color '#C00000'</code>
///   <para>Applies bold red styling to matching text.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeWordTextStyle")]
[Alias("WordTextStyle")]
[OutputType(typeof(WordParagraph))]
public sealed class SetOfficeWordTextStyleCommand : PSCmdlet
{
    /// <summary>Word text item to update.</summary>
    [Parameter(ValueFromPipeline = true, Position = 0)]
    public WordParagraph? InputObject { get; set; }

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
    public bool? Bold { get; set; }

    /// <summary>Set or clear italic formatting.</summary>
    [Parameter]
    public bool? Italic { get; set; }

    /// <summary>Underline style to apply.</summary>
    [Parameter]
    public WordUnderlineStyle? Underline { get; set; }

    /// <summary>Text color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? Color { get; set; }

    /// <summary>Font size in points.</summary>
    [Parameter]
    public double? FontSize { get; set; }

    /// <summary>Font family name.</summary>
    [Parameter]
    public string? FontFamily { get; set; }

    /// <summary>Highlight color.</summary>
    [Parameter]
    public WordHighlightColor? Highlight { get; set; }

    /// <summary>Set or clear strikethrough.</summary>
    [Parameter]
    public bool? Strike { get; set; }

    /// <summary>Set or clear double strikethrough.</summary>
    [Parameter]
    public bool? DoubleStrike { get; set; }

    /// <summary>Capitalization style.</summary>
    [Parameter]
    public WordCapsStyle? CapsStyle { get; set; }

    /// <summary>Character spacing in twentieths of a point.</summary>
    [Parameter]
    public int? Spacing { get; set; }

    /// <summary>Set or clear outline effect.</summary>
    [Parameter]
    public bool? Outline { get; set; }

    /// <summary>Set or clear shadow effect.</summary>
    [Parameter]
    public bool? Shadow { get; set; }

    /// <summary>Set or clear emboss effect.</summary>
    [Parameter]
    public bool? Emboss { get; set; }

    /// <summary>Emit the updated text item.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var text = InputObject ?? WordDslContext.Current?.CurrentParagraph;
        if (text == null)
        {
            throw new PSInvalidOperationException("Provide a Word text item or run this command inside a Word paragraph DSL scope.");
        }

        if (IsBound(nameof(Text))) text.Text = Text ?? string.Empty;
        if (Style.HasValue) text.SetCharacterStyle(Style.Value);
        if (!string.IsNullOrWhiteSpace(StyleId)) text.SetCharacterStyleId(StyleId!);
        if (IsBound(nameof(Bold))) text.Bold = Bold ?? false;
        if (IsBound(nameof(Italic))) text.Italic = Italic ?? false;
        if (Underline.HasValue) text.Underline = Underline.Value;
        if (IsBound(nameof(Color))) text.ColorHex = Color ?? string.Empty;
        if (IsBound(nameof(FontSize))) text.FontSizePoints = FontSize;
        if (IsBound(nameof(FontFamily))) text.FontFamily = FontFamily;
        if (Highlight.HasValue) text.Highlight = Highlight.Value;
        if (IsBound(nameof(Strike))) text.Strike = Strike ?? false;
        if (IsBound(nameof(DoubleStrike))) text.DoubleStrike = DoubleStrike ?? false;
        if (CapsStyle.HasValue) text.CapsStyle = CapsStyle.Value;
        if (IsBound(nameof(Spacing))) text.Spacing = Spacing;
        if (IsBound(nameof(Outline))) text.Outline = Outline ?? false;
        if (IsBound(nameof(Shadow))) text.Shadow = Shadow ?? false;
        if (IsBound(nameof(Emboss))) text.Emboss = Emboss ?? false;

        if (PassThru.IsPresent)
        {
            WriteObject(text);
        }
    }

    private bool IsBound(string parameterName) => MyInvocation.BoundParameters.ContainsKey(parameterName);
}
