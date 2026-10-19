using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Markdown;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Markdown;
using PSWriteOffice.Services.Table;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Adds a Markdown table from objects.</summary>
/// <example>
///   <summary>Add a table from input objects.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>MarkdownTable -InputObject $rows</code>
///   <para>Appends a Markdown table using the supplied objects.</para>
/// </example>
/// <example>
///   <summary>Append multiple tables to the same document.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$doc = New-OfficeMarkdown -Path .\Report.md -NoSave -PassThru
///   $doc | MarkdownTable -InputObject $summary -PassThru | MarkdownTable -InputObject $details</code>
///   <para>Creates two tables in sequence within the same Markdown document.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeMarkdownTable", DefaultParameterSetName = ParameterSetContext)]
[Alias("MarkdownTable")]
[OutputType(typeof(MarkdownDoc))]
public sealed class AddOfficeMarkdownTableCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPipelineDocument = "PipelineDocument";
    private readonly List<object?> _items = new();
    private MarkdownDoc? _document;

    /// <summary>Markdown document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetDocument)]
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetPipelineDocument)]
    public MarkdownDoc Document { get; set; } = null!;

    /// <summary>Objects to convert into a Markdown table.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = ParameterSetContext)]
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPipelineDocument)]
    public object? InputObject { get; set; }

    /// <summary>Projection to apply before writing the table.</summary>
    [Parameter]
    public OfficeTableView View { get; set; } = OfficeTableView.Normal;

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

    /// <summary>Emit the Markdown document after appending the table.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void BeginProcessing()
    {
        if (ParameterSetName == ParameterSetContext)
        {
            _document = ResolveDocument();
        }
    }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (ParameterSetName == ParameterSetPipelineDocument)
        {
            RenderTable(Document, BuildRows(InputObject));
            if (PassThru.IsPresent)
            {
                WriteObject(Document);
            }

            return;
        }

        AddInput(InputObject);
    }

    /// <inheritdoc />
    protected override void EndProcessing()
    {
        if (ParameterSetName == ParameterSetPipelineDocument)
        {
            return;
        }

        var doc = _document ?? ResolveDocument();
        RenderTable(doc, TableInputCollector.RequireRows(_items, nameof(InputObject)));
        if (PassThru.IsPresent)
        {
            WriteObject(doc);
        }
    }

    private void RenderTable(MarkdownDoc doc, object[] rows)
    {
        var projectedRows = TableViewProjection.Project(rows, View);
        var options = PowerShellObjectNormalizerOptions.ForTable(
            CollectionSeparator,
            DictionaryEntrySeparator,
            DictionaryKeyValueSeparator,
            MaxCollectionItems,
            MaxNestingDepth);
        var normalizedRows = PowerShellObjectNormalizer.NormalizeItems(projectedRows, options);
        if (DisableAutoAlign.IsPresent)
        {
            doc.TableFrom(normalizedRows);
        }
        else
        {
            doc.TableFromAuto(normalizedRows);
        }
    }

    private void AddInput(object? value)
    {
        TableInputCollector.AddInput(_items, value);
    }

    private static object[] BuildRows(object? value)
    {
        var items = new List<object?>();
        TableInputCollector.AddInput(items, value);
        return TableInputCollector.RequireRows(items, nameof(InputObject));
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
