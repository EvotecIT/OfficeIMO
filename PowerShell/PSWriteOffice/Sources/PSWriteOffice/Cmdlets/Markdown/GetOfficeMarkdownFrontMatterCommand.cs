using PSWriteOffice.Foundation;
using System.Management.Automation;
using OfficeIMO.Markdown;
using OfficeIMO.Yaml;
using PSWriteOffice.Services.Markdown;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Gets top-level YAML front matter entries from a Markdown document as Key, Value, Node and SourceSpan objects.</summary>
/// <example>
///   <summary>Read selected front matter from a Markdown file.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$metadata = Get-OfficeMarkdownFrontMatter -Path .\Report.md -Key 'title'
/// $metadata |
///     Select-Object -Property Key, Value |
///     Format-Table -AutoSize</code>
///   <para>Parses a Markdown file and returns matching front matter entries for metadata proof.</para>
/// </example>
/// <example>
///   <summary>Parse front matter from generated text.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$markdown = "---`ntitle: Report`nstatus: Ready`n---`n# Report"
/// Get-OfficeMarkdownFrontMatter -Text $markdown |
///     Select-Object -Property Key, Value</code>
///   <para>Parses Markdown text directly when the document has not been saved yet.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeMarkdownFrontMatter", DefaultParameterSetName = "Document")]
[OutputType(typeof(PSObject))]
public sealed class GetOfficeMarkdownFrontMatterCommand : PSCmdlet {
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

    /// <summary>Optional wildcard pattern matched against front matter keys.</summary>
    [Parameter]
    public string? Key { get; set; }

    /// <summary>Use case-sensitive matching for key filters.</summary>
    [Parameter]
    public SwitchParameter CaseSensitive { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var document = MarkdownDocumentResolver.Resolve(this, ParameterSetName, Document, Path, LiteralPath, Text, Options);

        var wildcardOptions = CaseSensitive
            ? WildcardOptions.None
            : WildcardOptions.IgnoreCase;
        var keyPattern = string.IsNullOrWhiteSpace(Key)
            ? null
            : new WildcardPattern(Key, wildcardOptions);

        foreach (var entry in document.FrontMatterEntries) {
            if (keyPattern != null && !keyPattern.IsMatch(entry.Key)) {
                continue;
            }

            var result = new PSObject();
            result.Properties.Add(new PSNoteProperty("Key", entry.Key));
            // Plain .NET values: strings, booleans, numbers, dictionaries for mappings and lists for sequences.
            result.Properties.Add(new PSNoteProperty("Value", YamlSerializer.Deserialize<object>(entry.Value)));
            result.Properties.Add(new PSNoteProperty("Node", entry.Value));
            result.Properties.Add(new PSNoteProperty("SourceSpan", entry.SourceSpan));
            WriteObject(result);
        }
    }
}