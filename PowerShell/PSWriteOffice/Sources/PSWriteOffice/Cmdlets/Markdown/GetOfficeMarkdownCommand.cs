using System.IO;
using System.Management.Automation;
using OfficeIMO.Markdown;
using PSWriteOffice.Services.Markdown;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Parses Markdown text or files into a Markdown document model.</summary>
/// <para>Returns an <see cref="MarkdownDoc"/> for inspection or further rendering.</para>
/// <example>
///   <summary>Parse a Markdown file.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$md = Get-OfficeMarkdown -Path .\README.md</code>
///   <para>Loads the file into a Markdown document object.</para>
/// </example>
/// <example>
///   <summary>Parse Markdown text in-memory.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$md = Get-OfficeMarkdown -Text '# Title`n`nBody text'</code>
///   <para>Parses Markdown text directly into a document model.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeMarkdown", DefaultParameterSetName = ParameterSetPath)]
[OutputType(typeof(MarkdownDoc))]
public sealed class GetOfficeMarkdownCommand : PSCmdlet
    , IMarkdownReaderOptionSource {
    private const string ParameterSetPath = "Path";
    private const string ParameterSetText = "Text";

    /// <summary>Path to the Markdown file.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Markdown text to parse.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetText)]
    public string Text { get; set; } = string.Empty;

    /// <summary>Optional reader options.</summary>
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

    MarkdownReaderOptions? IMarkdownReaderOptionSource.ReaderOptions => Options;

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = MarkdownOptionUtilities.BuildReaderOptions(this);

        MarkdownDoc document;
        if (ParameterSetName == ParameterSetPath) {
            var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
            if (!File.Exists(resolvedPath)) {
                throw new FileNotFoundException($"File '{resolvedPath}' was not found.", resolvedPath);
            }

            document = MarkdownDoc.Load(resolvedPath, options);
        } else {
            document = MarkdownReader.Parse(Text ?? string.Empty, options);
        }

        WriteObject(document);
    }
}