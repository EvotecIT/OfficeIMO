using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Markdown;
using PSWriteOffice.Services.Markdown;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Adds a Markdown task list.</summary>
/// <example>
///   <summary>Add a release checklist.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeMarkdown -Path .\ReleaseChecklist.md {
///     Add-OfficeMarkdownHeading -Level 1 -Text 'Release checklist'
///     Add-OfficeMarkdownTaskList -Items 'Generate examples', 'Refresh docs', 'Validate artifacts', 'Publish release' -Completed 0, 1
/// }</code>
///   <para>Appends an unordered task list and marks selected zero-based items as completed.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeMarkdownTaskList", DefaultParameterSetName = ParameterSetContext)]
[Alias("MarkdownTaskList")]
[OutputType(typeof(MarkdownDoc))]
public sealed class AddOfficeMarkdownTaskListCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>Markdown document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public MarkdownDoc Document { get; set; } = null!;

    /// <summary>Task text entries to include in the checklist.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string[] Items { get; set; } = Array.Empty<string>();

    /// <summary>Zero-based item indexes that should be marked complete.</summary>
    [Parameter]
    public int[] Completed { get; set; } = Array.Empty<int>();

    /// <summary>Mark every task as completed.</summary>
    [Parameter]
    public SwitchParameter AllCompleted { get; set; }

    /// <summary>Emit the updated Markdown document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var doc = ResolveDocument();
        var items = NormalizeItems(Items);
        if (items.Count == 0)
        {
            throw new PSArgumentException("Provide at least one task item.", nameof(Items));
        }

        var completed = BuildCompletedSet(items.Count);
        doc.Ul(builder =>
        {
            for (var i = 0; i < items.Count; i++)
            {
                builder.ItemTask(items[i], completed.Contains(i));
            }
        });

        if (PassThru.IsPresent)
        {
            WriteObject(doc);
        }
    }

    private HashSet<int> BuildCompletedSet(int itemCount)
    {
        if (AllCompleted.IsPresent)
        {
            return Enumerable.Range(0, itemCount).ToHashSet();
        }

        var completed = new HashSet<int>();
        foreach (var index in Completed ?? Array.Empty<int>())
        {
            if (index < 0 || index >= itemCount)
            {
                throw new PSArgumentOutOfRangeException(nameof(Completed), index, "Completed item indexes must match the provided Items collection.");
            }

            completed.Add(index);
        }

        return completed;
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
