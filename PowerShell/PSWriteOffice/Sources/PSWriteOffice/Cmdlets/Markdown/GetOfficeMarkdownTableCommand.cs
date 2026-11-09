using PSWriteOffice.Foundation;
using System.Management.Automation;
using OfficeIMO.Markdown;
using PSWriteOffice.Services.Markdown;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Gets Markdown tables from a Markdown document.</summary>
/// <example>
///   <summary>Read tables as PowerShell objects.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Open-OfficeMarkdown -Path .\Report.md | Get-OfficeMarkdownTable -AsObject</code>
///   <para>Returns table rows as PowerShell objects using the Markdown header row as property names.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeMarkdownTable", DefaultParameterSetName = "Document")]
[OutputType(typeof(TableBlock), typeof(PSObject))]
public sealed class GetOfficeMarkdownTableCommand : PSCmdlet {
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPath = "Path";
    private const string ParameterSetText = "Text";

    /// <summary>Caller-owned document; omission uses the enclosing Markdown document.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = ParameterSetDocument), ValidateNotNull]
    public MarkdownDocument? Document { get; set; }
    /// <summary>Input file with FileSystem wildcard expansion.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = ParameterSetPath)]
    [SupportsWildcards, ValidateNotNullOrEmpty]
    public string Path { get; set; } = string.Empty;
    /// <summary>Literal input file, including a FileInfo pipeline.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Alias("FullName", "PSPath"), ValidateNotNullOrEmpty]
    public string LiteralPath { get; set; } = string.Empty;
    /// <summary>Markdown source text.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetText), AllowEmptyString]
    public string Text { get; set; } = string.Empty;
    /// <summary>Native reader options and optional native-property hashtable overrides.</summary>
    [Parameter(ParameterSetName = ParameterSetPath), Parameter(ParameterSetName = "LiteralPath"), Parameter(ParameterSetName = ParameterSetText)]
    public object[]? Options { get; set; }

    /// <summary>Emit table rows as PowerShell objects instead of raw table blocks.</summary>
    [Parameter]
    public SwitchParameter AsObject { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var document = MarkdownDocumentResolver.Resolve(this, ParameterSetName, Document, Path, LiteralPath, Text, Options);

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