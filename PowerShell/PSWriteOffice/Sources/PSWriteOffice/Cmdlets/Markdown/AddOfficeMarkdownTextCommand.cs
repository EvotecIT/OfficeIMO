using System.Linq;
using System.Management.Automation;
using OfficeIMO.Markdown;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Adds styled Markdown text to a paragraph or list item content block.</summary>
/// <example><summary>Write mixed emphasis.</summary>
/// <code>MarkdownParagraph { MarkdownText 'Status: '; MarkdownText 'Healthy' -Bold -Italic }</code></example>
[Cmdlet(VerbsCommon.Add, "OfficeMarkdownText", DefaultParameterSetName = "Default")]
[OutputType(typeof(void))]
[Alias("MarkdownText") ]
public sealed class AddOfficeMarkdownTextCommand : PSCmdlet {
    /// <summary>Literal text to append.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Default"), AllowEmptyString] public string Text { get; set; } = string.Empty;
    /// <summary>Apply bold emphasis.</summary>
    [Parameter] public SwitchParameter Bold { get; set; }
    /// <summary>Apply italic emphasis.</summary>
    [Parameter] public SwitchParameter Italic { get; set; }
    /// <summary>Apply strikethrough.</summary>
    [Parameter] public SwitchParameter Strikethrough { get; set; }
    /// <summary>Render literal inline code.</summary>
    [Parameter] public SwitchParameter Code { get; set; }
    /// <summary>Optional hyperlink URL.</summary>
    [Parameter] public string? Link { get; set; }
    /// <inheritdoc />
    protected override void ProcessRecord() {
        var target = OfficeDocumentContext.Find<InlineSequence>()
            ?? throw new PSInvalidOperationException("Use MarkdownText inside MarkdownParagraph or MarkdownListItem content.");
        var content = new InlineSequence();
        if (Link != null) content.Link(Text, Link);
        else if (Code) content.Code(Text);
        else content.Text(Text);
        if (Bold && Italic) content.ReplaceItems(new[] { new BoldItalicSequenceInline(Copy(content)) });
        else if (Bold) content.ReplaceItems(new[] { new BoldSequenceInline(Copy(content)) });
        else if (Italic) content.ReplaceItems(new[] { new ItalicSequenceInline(Copy(content)) });
        if (Strikethrough) content.ReplaceItems(new[] { new StrikethroughSequenceInline(Copy(content)) });
        target.ReplaceItems(target.Nodes.Concat(content.Nodes));
    }
    private static InlineSequence Copy(InlineSequence source) {
        var result = new InlineSequence();
        result.ReplaceItems(source.Nodes);
        return result;
    }
}
