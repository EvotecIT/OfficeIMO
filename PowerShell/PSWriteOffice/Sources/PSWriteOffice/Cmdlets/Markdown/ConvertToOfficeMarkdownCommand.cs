using System;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO.Markdown;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Converts objects into a Markdown table.</summary>
/// <para>Returns Markdown text by default; use <c>-PassThru</c> to emit a <see cref="MarkdownDoc"/>.</para>
/// <example>
///   <summary>Convert objects to Markdown table.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$markdown = $data | ConvertTo-OfficeMarkdown</code>
///   <para>Generates Markdown table text from the input objects.</para>
/// </example>
/// <example>
///   <summary>Emit a Markdown document for further editing.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$doc = $data | ConvertTo-OfficeMarkdown -PassThru
///   $doc.P('Totals above'); $doc.ToMarkdown()</code>
///   <para>Builds a table and appends more content using the MarkdownDoc API.</para>
/// </example>
/// <example>
///   <summary>Disable auto alignment.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$markdown = $data | ConvertTo-OfficeMarkdown -DisableAutoAlign</code>
///   <para>Forces left-aligned columns instead of auto-aligned output.</para>
/// </example>
[Cmdlet(VerbsData.ConvertTo, "OfficeMarkdown")]
[OutputType(typeof(string), typeof(MarkdownDoc))]
public sealed class ConvertToOfficeMarkdownCommand : PSCmdlet
{
    private readonly List<object?> _items = new();

    /// <summary>Objects to convert into Markdown.</summary>
    [Parameter(ValueFromPipeline = true)]
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
    public SwitchParameter DisableAutoAlign { get; set; }

    /// <summary>Emit a Markdown document object instead of text.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

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
        var doc = MarkdownDoc.Create();
        if (DisableAutoAlign.IsPresent)
        {
            doc.TableFrom(normalizedItems);
        }
        else
        {
            doc.TableFromAuto(normalizedItems);
        }

        if (PassThru.IsPresent)
        {
            WriteObject(doc);
        }
        else
        {
            WriteObject(doc.ToMarkdown());
        }
    }

}
