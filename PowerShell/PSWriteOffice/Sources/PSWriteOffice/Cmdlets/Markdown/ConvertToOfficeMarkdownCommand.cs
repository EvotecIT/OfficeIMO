using System;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO.Markdown;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Converts objects into a Markdown table.</summary>
/// <para>Returns Markdown text by default; use <c>-AsDocument</c> to emit a <see cref="MarkdownDocument"/>.</para>
/// <example>
///   <summary>Convert objects to Markdown table.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$markdown = $data | ConvertTo-OfficeMarkdown</code>
///   <para>Generates Markdown table text from the input objects.</para>
/// </example>
/// <example>
///   <summary>Emit a Markdown document for further editing.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$doc = $data | ConvertTo-OfficeMarkdown -AsDocument
///   $doc.P('Totals above'); $doc.ToMarkdown()</code>
///   <para>Builds a table and appends more content using the MarkdownDocument API.</para>
/// </example>
/// <example>
///   <summary>Disable auto alignment.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$markdown = $data | ConvertTo-OfficeMarkdown -NoAutoAlign</code>
///   <para>Forces left-aligned columns instead of auto-aligned output.</para>
/// </example>
[Cmdlet(VerbsData.ConvertTo, "OfficeMarkdown", DefaultParameterSetName = "Default")]
[OutputType(typeof(string), typeof(MarkdownDocument))]
public sealed class ConvertToOfficeMarkdownCommand : PSCmdlet
{
    private readonly List<object?> _items = new();

    /// <summary>Objects to convert into Markdown.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = "Default")]
    public object? InputObject { get; set; }

    /// <summary>Text used between items when a cell contains a collection.</summary>
    [Parameter]
    [AllowEmptyString]
    public string CollectionSeparator { get; set; } = ", ";

    /// <summary>Text used between entries when a cell contains a dictionary.</summary>
    [Parameter]
    [AllowEmptyString]
    public string DictionaryEntrySeparator { get; set; } = "; ";

    /// <summary>Text used between a dictionary key and value.</summary>
    [Parameter]
    [AllowEmptyString]
    public string DictionaryKeyValueSeparator { get; set; } = ": ";

    /// <summary>Maximum number of items allowed in one nested collection or dictionary cell. Defaults to 1,048,575; increase explicitly for trusted larger values.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int MaxCollectionItems { get; set; } = PowerShellObjectNormalizerOptions.DefaultMaxCollectionItems;

    /// <summary>Maximum nesting depth allowed while normalizing one cell value. Defaults to 64; increase explicitly for trusted deeper values.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int MaxNestingDepth { get; set; } = PowerShellObjectNormalizerOptions.DefaultMaxNestingDepth;

    /// <summary>Disable automatic alignment heuristics for tables.</summary>
    [Parameter]
    public SwitchParameter NoAutoAlign { get; set; }

    /// <summary>Emit a Markdown document object instead of text.</summary>
    [Parameter]
    public SwitchParameter AsDocument { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        _items.Add(InputObject);
    }

    /// <inheritdoc />
    protected override void EndProcessing()
    {
        if (_items.Count == 0)
        {
            return;
        }

        var options = PowerShellObjectNormalizerOptions.ForTable(
            CollectionSeparator,
            DictionaryEntrySeparator,
            DictionaryKeyValueSeparator,
            MaxCollectionItems,
            MaxNestingDepth);
        var normalizedItems = PowerShellObjectNormalizer.NormalizeItems(_items, options);
        var doc = MarkdownDocument.Create();
        if (NoAutoAlign.IsPresent)
        {
            doc.AddTable((object)normalizedItems);
        }
        else
        {
            doc.AddTable((object)normalizedItems, new MarkdownTableOptions());
        }

        if (AsDocument.IsPresent)
        {
            WriteObject(doc);
        }
        else
        {
            WriteObject(doc.ToMarkdown());
        }
    }

}
