using PSWriteOffice.Foundation;
using System.Management.Automation;
using OfficeIMO.Markdown;
using PSWriteOffice.Services.Markdown;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Gets heading metadata from a Markdown document.</summary>
/// <para>Returns heading level, text, resolved anchor, and the backing heading block.</para>
/// <example>
///   <summary>Read headings from Markdown text.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficeMarkdownHeading -Text "# Title`n`n## Details"</code>
///   <para>Parses Markdown text and returns the document headings.</para>
/// </example>
/// <example>
///   <summary>Inspect a parsed Markdown document.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Open-OfficeMarkdown -Path .\README.md | Get-OfficeMarkdownHeading -MinLevel 2</code>
///   <para>Returns headings from an existing Markdown document object.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeMarkdownHeading", DefaultParameterSetName = "Document")]
[OutputType(typeof(MarkdownDocument.HeadingInfo))]
public sealed class GetOfficeMarkdownHeadingCommand : PSCmdlet {
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

    /// <summary>Minimum heading level to return.</summary>
    [Parameter]
    [ValidateRange(1, 6)]
    public int MinLevel { get; set; } = 1;

    /// <summary>Maximum heading level to return.</summary>
    [Parameter]
    [ValidateRange(1, 6)]
    public int MaxLevel { get; set; } = 6;

    /// <summary>Optional wildcard pattern matched against heading text.</summary>
    [Parameter]
    public string? HeadingText { get; set; }

    /// <summary>Optional wildcard pattern matched against resolved heading anchors.</summary>
    [Parameter]
    public string? Anchor { get; set; }

    /// <summary>Use case-sensitive matching for text and anchor filters.</summary>
    [Parameter]
    public SwitchParameter CaseSensitive { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (MinLevel > MaxLevel) {
            throw new PSArgumentException("-MinLevel cannot be greater than -MaxLevel.");
        }

        var document = MarkdownDocumentResolver.Resolve(this, ParameterSetName, Document, Path, LiteralPath, Text, Options);

        var wildcardOptions = CaseSensitive
            ? WildcardOptions.None
            : WildcardOptions.IgnoreCase;
        var textPattern = string.IsNullOrWhiteSpace(HeadingText)
            ? null
            : new WildcardPattern(HeadingText, wildcardOptions);
        var anchorPattern = string.IsNullOrWhiteSpace(Anchor)
            ? null
            : new WildcardPattern(Anchor!.TrimStart('#'), wildcardOptions);

        foreach (var heading in document.GetHeadingInfos()) {
            if (heading.Level < MinLevel || heading.Level > MaxLevel) {
                continue;
            }

            if (textPattern != null && !textPattern.IsMatch(heading.Text)) {
                continue;
            }

            if (anchorPattern != null && !anchorPattern.IsMatch(heading.Anchor)) {
                continue;
            }

            WriteObject(heading);
        }
    }
}