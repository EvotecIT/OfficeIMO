using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Markdown;
using PSWriteOffice.Services.Markdown;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Adds a Markdown list.</summary>
/// <example>
///   <summary>Add ordered and unordered lists to a report.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeMarkdown -Path .\Plan.md {
///     Add-OfficeMarkdownHeading -Level 2 -Text 'Priorities'
///     Add-OfficeMarkdownList -Items 'PDF examples', 'PowerPoint layout', 'Word report polish'
///     Add-OfficeMarkdownHeading -Level 2 -Text 'Release order'
///     Add-OfficeMarkdownList -Items 'Validate artifacts', 'Generate docs', 'Publish release' -Ordered -Start 1
/// }</code>
///   <para>Appends bullet and ordered lists to the current Markdown document.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeMarkdownList", DefaultParameterSetName = ParameterSetContext)]
[Alias("MarkdownList")]
[OutputType(typeof(MarkdownDoc))]
public sealed class AddOfficeMarkdownListCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>Markdown document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public MarkdownDoc Document { get; set; } = null!;

    /// <summary>List items to add.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string[] Items { get; set; } = Array.Empty<string>();

    /// <summary>Use an ordered list instead of bullets.</summary>
    [Parameter]
    public SwitchParameter Ordered { get; set; }

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
        var items = NormalizeItems(Items);
        if (items.Count == 0)
        {
            throw new PSArgumentException("Provide at least one list item.", nameof(Items));
        }

        if (Ordered.IsPresent)
        {
            doc.Ol(items, Start);
        }
        else
        {
            doc.Ul(items);
        }

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

    private MarkdownDoc ResolveDocument()
    {
        if (ParameterSetName == ParameterSetDocument)
        {
            return Document ?? throw new PSArgumentException("Provide a Markdown document.");
        }

        return MarkdownDslContext.Require(this).Document;
    }
}
