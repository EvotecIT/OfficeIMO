using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO.Markdown;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Adds a plain or rich item to a MarkdownList block.</summary>
/// <example><summary>Add styled list content.</summary>
/// <code>MarkdownList { MarkdownListItem { MarkdownText 'Review' -Bold } }</code></example>
[Cmdlet(VerbsCommon.Add, "OfficeMarkdownListItem", DefaultParameterSetName = "Text")]
[OutputType(typeof(void))]
[Alias("MarkdownListItem") ]
public sealed class AddOfficeMarkdownListItemCommand : PSCmdlet {
    /// <summary>Plain item text.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Text"), AllowEmptyString]
    public string Text { get; set; } = string.Empty;
    /// <summary>Styled MarkdownText content.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Content")]
    public ScriptBlock? Content { get; set; }
    /// <inheritdoc />
    protected override void ProcessRecord() {
        var items = OfficeDocumentContext.Find<List<ListItem>>()
            ?? throw new PSInvalidOperationException("Use MarkdownListItem inside MarkdownList.");
        var inline = new InlineSequence();
        var item = new ListItem(inline);
        if (Content != null) { using (OfficeDocumentContext.Push(item))
            using (OfficeDocumentContext.Push(inline)) Content.InvokeReturnAsIs(); }
        else inline.Text(Text);
        items.Add(item);
    }
}
