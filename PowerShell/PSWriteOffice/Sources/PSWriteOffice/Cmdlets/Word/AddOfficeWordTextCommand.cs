using OfficeIMO.Drawing;
using System;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Word;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Adds styled text or a hyperlink to the current or given Word paragraph.</summary>
/// <para>A bare Size means points; unit-suffixed lengths are accepted. Hyperlinks use the Hyperlink style unless NoLinkStyle is supplied.</para>
/// <example><summary>Append styled text.</summary>
/// <code>WordParagraph { WordText 'Important: ' -Bold -Font Aptos -Size 11pt; WordText 'read first' -Italic }</code></example>
/// <example><summary>Add a link and a bookmark reference.</summary>
/// <code>WordParagraph { WordText 'Example' -Link 'https://example.org' -Tooltip 'External reference'; WordText 'Summary' -Anchor Summary }</code></example>
[Cmdlet(VerbsCommon.Add, "OfficeWordText", DefaultParameterSetName = "Text")]
[OutputType(typeof(WordParagraph), typeof(WordHyperLink))]
[Alias("WordText") ]
public sealed class AddOfficeWordTextCommand : PSCmdlet {
    /// <summary>Text segments to append.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Text")]
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Link")]
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Anchor")]
    public string[] Text { get; set; } = Array.Empty<string>();

    /// <summary>Paragraph to receive text outside a content block.</summary>
    [Parameter(ValueFromPipeline = true)]
    public WordParagraph? Paragraph { get; set; }

    /// <summary>Apply bold formatting.</summary>
    [Parameter]
    public SwitchParameter Bold { get; set; }

    /// <summary>Apply italic formatting.</summary>
    [Parameter]
    public SwitchParameter Italic { get; set; }

    /// <summary>Apply strikethrough.</summary>
    [Parameter]
    public SwitchParameter Strikethrough { get; set; }

    /// <summary>Underline style.</summary>
    [Parameter]
    public WordUnderlineStyle? Underline { get; set; }

    /// <summary>Text color, including a semantic theme role.</summary>
    [Parameter, OfficeColorTransformation, ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? Color { get; set; }

    /// <summary>Font family.</summary>
    [Parameter]
    public string? Font { get; set; }

    /// <summary>Font size; a bare number means points.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? Size { get; set; }

    /// <summary>Absolute hyperlink URL.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "Link")]
    public Uri? Link { get; set; }

    /// <summary>Bookmark target in this document.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "Anchor"), ValidateNotNullOrEmpty]
    public string? Anchor { get; set; }

    /// <summary>Hyperlink tooltip.</summary>
    [Parameter(ParameterSetName = "Link"), Parameter(ParameterSetName = "Anchor")]
    public string? Tooltip { get; set; }

    /// <summary>Disable hyperlink navigation history.</summary>
    [Parameter(ParameterSetName = "Link"), Parameter(ParameterSetName = "Anchor")]
    public SwitchParameter NoHistory { get; set; }

    /// <summary>Do not apply the built-in Hyperlink character style.</summary>
    [Parameter(ParameterSetName = "Link"), Parameter(ParameterSetName = "Anchor")]
    public SwitchParameter NoLinkStyle { get; set; }

    /// <summary>Run highlight or shading color; explicit null clears both.</summary>
    [Parameter, AllowNull, OfficeColorTransformation(AllowNull = true)]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? BackgroundColor { get; set; }
    /// <summary>Normal, superscript or subscript baseline.</summary>
    [Parameter]
    public OfficeTextBaseline? Baseline { get; set; }

    /// <summary>Emit each newly created run, or each hyperlink for Link and Anchor.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void BeginProcessing() {
        if (Link != null && !Link.IsAbsoluteUri)
            throw new PSArgumentException("Provide an absolute URL such as https://example.org or mailto:user@example.org.", nameof(Link));
        if (ParameterSetName == "Anchor" && string.IsNullOrWhiteSpace(Anchor))
            throw new PSArgumentException("Provide a bookmark anchor.", nameof(Anchor));
        // Validate before appending any text or creating an implicit paragraph.
        if (Size.HasValue && (Size.Value.ToPoints() < 0 || Size.Value.ToPoints() > int.MaxValue / 2D))
            throw new ArgumentOutOfRangeException(nameof(Size), "The Word font size is outside the supported range.");
    }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var applyStyle = CaptureStyle();
        var recording = Paragraph == null ? OfficeDocumentContext.Find<WordCellContent>() : null;
        if (recording != null) {
            if (PassThru) throw new PSArgumentException("Recorded cell content emits native runs only when attached to a table; omit PassThru inside this factory.");
            var link = Link;
            var anchor = Anchor;
            bool useLinkStyle = !NoLinkStyle;
            bool history = !NoHistory;
            string tooltip = Tooltip ?? string.Empty;
            foreach (string value in Text) {
                string text = value;
                recording.Text.Add(target => {
                    WordHyperLink? hyperlink = link != null ? target.AddHyperLink(text, link, useLinkStyle, tooltip, history)
                        : anchor != null ? target.AddHyperLink(text, anchor, useLinkStyle, tooltip, history) : null;
                    if (hyperlink != null) { foreach (var run in hyperlink.Runs) applyStyle(run); }
                    else applyStyle(target.AddFormattedText(text));
                });
            }
            return;
        }
        var paragraph = Paragraph ?? OfficeDocumentContext.Find<WordParagraph>();
        if (paragraph == null) {
            var context = WordDslContext.Require(this);
            paragraph = context.CurrentParagraph ?? context.RequireParagraphHost().AddParagraph();
        }
        foreach (string entry in Text) {
            WordHyperLink? hyperlink = null;
            if (ParameterSetName == "Link")
                hyperlink = paragraph.AddHyperLink(entry, Link!, !NoLinkStyle, Tooltip ?? string.Empty, !NoHistory);
            else if (ParameterSetName == "Anchor")
                hyperlink = paragraph.AddHyperLink(entry, Anchor!, !NoLinkStyle, Tooltip ?? string.Empty, !NoHistory);

            if (hyperlink != null) {
                foreach (var run in hyperlink.Runs) applyStyle(run);
                if (PassThru) WriteObject(hyperlink);
            } else {
                var run = paragraph.AddFormattedText(entry);
                applyStyle(run);
                if (PassThru) WriteObject(run);
            }
        }
    }

    private Action<WordParagraph> CaptureStyle() {
        var bold = OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(Bold));
        var italic = OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(Italic));
        var strike = OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(Strikethrough));
        var underline = Underline;
        var color = Color;
        var font = Font;
        var size = Size;
        var background = BackgroundColor;
        bool backgroundBound = MyInvocation.BoundParameters.ContainsKey(nameof(BackgroundColor));
        var baseline = Baseline;
        return run => {
        // Preserve theme hyperlink defaults unless the caller supplies an override.
        if (bold.HasValue) run.Bold = bold.Value;
        if (italic.HasValue) run.Italic = italic.Value;
        if (strike.HasValue) run.Strikethrough = strike.Value;
        if (underline.HasValue) run.SetUnderline(underline.Value);
        if (size.HasValue) run.FontSize = size;
        if (!string.IsNullOrWhiteSpace(font)) run.SetFontFamily(font!);
        if (color.HasValue) run.Color = color;
        if (backgroundBound) {
            if (background.HasValue && background.Value.IsThemeRole) {
                run.BackgroundColor = OfficeColor.ParseHex(WordRunColor.Create(background.Value, run._document).Val!.Value!);
            } else if (background.HasValue) WordTextRunService.ApplyBackground(run, background.Value.ToString());
            else { run.Highlight = null; run.BackgroundColor = null; }
        }
        if (baseline.HasValue) run.VerticalTextAlignment = baseline.Value switch {
            OfficeTextBaseline.Superscript => WordVerticalTextPosition.Superscript,
            OfficeTextBaseline.Subscript => WordVerticalTextPosition.Subscript,
            _ => WordVerticalTextPosition.Baseline
        };
    };
    }
}
