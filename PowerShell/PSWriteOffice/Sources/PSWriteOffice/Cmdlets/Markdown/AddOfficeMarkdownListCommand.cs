using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Markdown;
using PSWriteOffice.Services.Markdown;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Adds a Markdown list.</summary>
/// <example>
///   <summary>Add ordered and unordered lists to a report.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeMarkdown -Path .\Plan.md {
///     Add-OfficeMarkdownHeading -Level 2 -Text 'Priorities'
///     Add-OfficeMarkdownList -Items 'PDF examples', 'PowerPoint layout', 'Word report polish'
///     Add-OfficeMarkdownHeading -Level 2 -Text 'Release order'
///     Add-OfficeMarkdownList -Items 'Validate artifacts', 'Generate docs', 'Publish release' -Numbered -Start 1
/// }</code>
///   <para>Appends bullet and ordered lists to the current Markdown document.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeMarkdownList", DefaultParameterSetName = "Items")]
[OutputType(typeof(MarkdownDocument))]
[Alias("MarkdownList") ]
public sealed class AddOfficeMarkdownListCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>Markdown document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "DocumentItems")]
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "DocumentContent")]
    public MarkdownDocument Document { get; set; } = null!;

    /// <summary>List items to add.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Items")]
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "DocumentItems")]
    public string[] Items { get; set; } = Array.Empty<string>();

    /// <summary>List items composed with MarkdownListItem children.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Content")]
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "DocumentContent")]
    public ScriptBlock? Content { get; set; }

    /// <summary>Use an ordered list instead of bullets.</summary>
    [Parameter]
    public SwitchParameter Numbered { get; set; }

    /// <summary>Starting number for ordered lists.</summary>
    [Parameter]
    public int Start { get; set; } = 1;

    /// <summary>Emit the Markdown document after appending the list.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var doc = ResolveDocument();
        var items = new List<ListItem>();
        if (Content != null) {
            using (OfficeDocumentContext.Push(items)) Content.InvokeReturnAsIs();
        } else items.AddRange(NormalizeItems(Items).Select(ListItem.Text));
        IMarkdownBlock list;
        if (Numbered) { var block = new OrderedListBlock { Start = Start }; block.Items.AddRange(items); list = block; }
        else { var block = new UnorderedListBlock(); block.Items.AddRange(items); list = block; }
        var parentItem = ParameterSetName.StartsWith("Document", StringComparison.Ordinal) ? null : OfficeDocumentContext.Find<ListItem>();
        if (parentItem != null) parentItem.NestedBlocks.Add(list);
        else doc.Add(list);

        if (PassThru.IsPresent)
        {
            WriteObject(doc);
        }
    }

    private static List<string> NormalizeItems(IEnumerable<string>? items)
    {
        return items == null
            ? new List<string>()
            : items.Select(item => item?.Trim())
                .Where(item => !string.IsNullOrWhiteSpace(item))
                .Cast<string>()
                .ToList();
    }

    private MarkdownDocument ResolveDocument()
    {
        if (ParameterSetName.StartsWith("Document", StringComparison.Ordinal))
        {
            return Document ?? throw new PSArgumentException("Provide a Markdown document.");
        }

        return MarkdownDslContext.Require(this).Document;
    }
}
