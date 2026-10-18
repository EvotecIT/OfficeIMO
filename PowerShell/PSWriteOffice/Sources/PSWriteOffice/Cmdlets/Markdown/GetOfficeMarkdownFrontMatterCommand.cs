using System.Management.Automation;
using OfficeIMO.Markdown;
using PSWriteOffice.Services.Markdown;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Gets YAML front matter entries from a Markdown document.</summary>
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
[Cmdlet(VerbsCommon.Get, "OfficeMarkdownFrontMatter", DefaultParameterSetName = ParameterSetPath)]
[OutputType(typeof(FrontMatterBlock.Entry))]
public sealed class GetOfficeMarkdownFrontMatterCommand : PSCmdlet
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

    /// <summary>Optional wildcard pattern matched against front matter keys.</summary>
    [Parameter]
    public string? Key { get; set; }

    MarkdownReaderOptions? IMarkdownReaderOptionSource.ReaderOptions => Options;

    /// <summary>Use case-sensitive matching for key filters.</summary>
    [Parameter]
    public SwitchParameter CaseSensitive { get; set; }

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

            WriteObject(entry);
        }
    }
}