using System;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO.Markdown;
using PSWriteOffice.Services.Markdown;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Gets Markdown tables from a Markdown document.</summary>
/// <example>
///   <summary>Read tables as PowerShell objects.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficeMarkdown -Path .\Report.md | Get-OfficeMarkdownTable -AsObject</code>
///   <para>Returns table rows as PowerShell objects using the Markdown header row as property names.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeMarkdownTable", DefaultParameterSetName = ParameterSetPath)]
[OutputType(typeof(TableBlock), typeof(PSObject))]
public sealed class GetOfficeMarkdownTableCommand : PSCmdlet
    , IMarkdownReaderOptionSource {
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPath = "Path";
    private const string ParameterSetText = "Text";

    /// <summary>Markdown document to inspect.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public MarkdownDoc Document { get; set; } = null!;

    /// <summary>Path to the Markdown file.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Markdown text to parse.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetText)]
    public string Text { get; set; } = string.Empty;

    /// <summary>Optional reader options used when parsing path or text input.</summary>
    [Parameter]
    [Alias("ReaderOptions")]
    public MarkdownReaderOptions? Options { get; set; }

    /// <summary>Named reader profile used when <see cref="Options"/> is not supplied.</summary>
    [Parameter]
    public MarkdownReaderOptions.MarkdownDialectProfile? Profile { get; set; }

    /// <summary>Base URI used to resolve and restrict relative Markdown links and images.</summary>
    [Parameter]
    public string? BaseUri { get; set; }

    /// <summary>Maximum Markdown input length accepted by the reader.</summary>
    [Parameter]
    public int? MaxInputCharacters { get; set; }

    /// <summary>Applies a built-in Markdown input normalization preset before parsing.</summary>
    [Parameter]
    public MarkdownInputNormalizationPreset? NormalizeInput { get; set; }

    /// <summary>Block file URLs while parsing Markdown links and images.</summary>
    [Parameter]
    public bool? DisallowFileUrls { get; set; }

    /// <summary>Allow data URLs while parsing Markdown links and images.</summary>
    [Parameter]
    public bool? AllowDataUrls { get; set; }

    /// <summary>Allow mailto URLs while parsing Markdown links.</summary>
    [Parameter]
    public bool? AllowMailtoUrls { get; set; }

    /// <summary>Allow protocol-relative URLs while parsing Markdown links and images.</summary>
    [Parameter]
    public bool? AllowProtocolRelativeUrls { get; set; }

    /// <summary>Restrict parsed URL schemes to the allow-list.</summary>
    [Parameter]
    public bool? RestrictUrlSchemes { get; set; }

    /// <summary>Allowed URL schemes when URL scheme restriction is enabled.</summary>
    [Parameter]
    public string[]? AllowedUrlScheme { get; set; }

    /// <summary>Emit table rows as PowerShell objects instead of raw table blocks.</summary>
    [Parameter]
    public SwitchParameter AsObject { get; set; }

    MarkdownReaderOptions? IMarkdownReaderOptionSource.ReaderOptions => Options;

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var document = MarkdownDocumentResolver.Resolve(
            this,
            ParameterSetName,
            ParameterSetDocument,
            Document,
            Path,
            Text,
            this);

        foreach (var table in document.DescendantTables()) {
            if (!AsObject) {
                WriteObject(table);
                continue;
            }

            foreach (var row in ConvertTableRows(table)) {
                WriteObject(row);
            }
        }
    }

    private static IEnumerable<PSObject> ConvertTableRows(TableBlock table) {
        var columnNames = GetColumnNames(table);

        foreach (var row in table.Rows) {
            var item = new PSObject();
            for (var i = 0; i < columnNames.Count; i++) {
                var value = i < row.Count ? row[i] : string.Empty;
                item.Properties.Add(new PSNoteProperty(columnNames[i], value));
            }

            yield return item;
        }
    }

    private static IReadOnlyList<string> GetColumnNames(TableBlock table) {
        var columnCount = table.Headers.Count;
        foreach (var row in table.Rows) {
            if (row.Count > columnCount) {
                columnCount = row.Count;
            }
        }

        var names = new List<string>(columnCount);
        var seen = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < columnCount; i++) {
            var name = i < table.Headers.Count ? table.Headers[i] : null;
            if (string.IsNullOrWhiteSpace(name)) {
                name = "Column" + (i + 1).ToString(System.Globalization.CultureInfo.InvariantCulture);
            }

            var baseName = name!.Trim();
            if (seen.TryGetValue(baseName, out var duplicateCount)) {
                duplicateCount++;
                seen[baseName] = duplicateCount;
                baseName += duplicateCount.ToString(System.Globalization.CultureInfo.InvariantCulture);
            } else {
                seen[baseName] = 1;
            }

            names.Add(baseName);
        }

        return names;
    }
}