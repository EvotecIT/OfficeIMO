using System;
using System.IO;
using System.Management.Automation;
using System.Text;
using OfficeIMO.Html;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;
using PSWriteOffice.Services.Markdown;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Converts HTML content to Markdown.</summary>
/// <para>Returns Markdown text or saves it to a file when <c>-OutputPath</c> is specified.</para>
/// <example>
///   <summary>Convert an HTML fragment to Markdown.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$markdown = ConvertFrom-OfficeMarkdownHtml -Html '&lt;h1&gt;Report&lt;/h1&gt;&lt;p&gt;Ready&lt;/p&gt;'</code>
///   <para>Returns Markdown text converted from the supplied HTML.</para>
/// </example>
/// <example>
///   <summary>Convert an HTML file to a Markdown document object.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$doc = ConvertFrom-OfficeMarkdownHtml -Path .\report.html -AsDocument</code>
///   <para>Returns a Markdown document for further editing or rendering.</para>
/// </example>
[Cmdlet(VerbsData.ConvertFrom, "OfficeMarkdownHtml", DefaultParameterSetName = ParameterSetHtml, SupportsShouldProcess = true)]
[Alias("ConvertFrom-MarkdownHtml")]
[OutputType(typeof(string), typeof(FileInfo), typeof(MarkdownDoc))]
public sealed class ConvertFromOfficeMarkdownHtmlCommand : PSCmdlet
    , IMarkdownWriteOptionSource {
    private const string ParameterSetHtml = "Html";
    private const string ParameterSetPath = "Path";

    /// <summary>HTML markup to convert.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = ParameterSetHtml)]
    public string Html { get; set; } = string.Empty;

    /// <summary>Path to an HTML file.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Optional output path for the Markdown file.</summary>
    [Parameter]
    [Alias("OutPath")]
    public string? OutputPath { get; set; }

    /// <summary>Emit a Markdown document object instead of Markdown text.</summary>
    [Parameter]
    public SwitchParameter AsDocument { get; set; }

    /// <summary>Emit a <see cref="FileInfo"/> when saving to disk.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <summary>Optional conversion options.</summary>
    [Parameter]
    public HtmlToMarkdownOptions? Options { get; set; }

    /// <summary>Use portable Markdown output when <see cref="Options"/> is not supplied.</summary>
    [Parameter]
    public SwitchParameter Portable { get; set; }

    /// <summary>Base URI used to resolve relative links and image sources.</summary>
    [Parameter]
    public string? BaseUri { get; set; }

    /// <summary>Convert the full HTML document instead of only body contents.</summary>
    [Parameter]
    public SwitchParameter IncludeDocumentChrome { get; set; }

    /// <summary>Preserve script, style, noscript, and template elements.</summary>
    [Parameter]
    public SwitchParameter PreserveScriptsAndStyles { get; set; }

    /// <summary>Drop unsupported block HTML instead of preserving it as raw HTML.</summary>
    [Parameter]
    public SwitchParameter DropUnsupportedBlocks { get; set; }

    /// <summary>Drop unsupported inline HTML instead of preserving it as raw HTML.</summary>
    [Parameter]
    public SwitchParameter DropUnsupportedInlineHtml { get; set; }

    /// <summary>Maximum input length, in characters, accepted by the converter.</summary>
    [Parameter]
    public int? MaxInputCharacters { get; set; }

    /// <summary>Controls how base64 data URI images are converted.</summary>
    [Parameter]
    public HtmlBase64ImageHandling? Base64ImageHandling { get; set; }

    /// <summary>Output directory for decoded base64 images when saving them to files.</summary>
    [Parameter]
    public string? Base64ImageOutputDirectory { get; set; }

    /// <summary>Controls whether repeated listing-card metadata is preserved or suppressed.</summary>
    [Parameter]
    public HtmlListingCardMetadataMode? ListingCardMetadataMode { get; set; }

    /// <summary>Maximum logical columns produced by expanding HTML table spans.</summary>
    [Parameter]
    public int? MaxTableExpandedColumns { get; set; }

    /// <summary>Optional Markdown writer options for generated Markdown text.</summary>
    [Parameter]
    public MarkdownWriteOptions? WriteOptions { get; set; }

    /// <summary>Friendly Markdown writer profile for generated Markdown text.</summary>
    [Parameter]
    public OfficeMarkdownWriteProfile? WriteProfile { get; set; }

    /// <summary>Controls how generated Markdown images are serialized.</summary>
    [Parameter]
    public MarkdownImageRenderingMode? ImageRenderingMode { get; set; }

    /// <summary>Markdown line ending: CRLF, LF, CR, or a literal line ending string.</summary>
    [Parameter]
    [ArgumentCompleter(typeof(OfficeLineEndingArgumentCompleter))]
    public string? LineEnding { get; set; }

    /// <summary>Unordered list marker: '-', '*', or '+'.</summary>
    [Parameter]
    public string? UnorderedListMarker { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        try {
            if (Options != null && Portable.IsPresent) {
                throw new PSArgumentException("Specify either -Options or -Portable, not both.");
            }

            if (AsDocument.IsPresent && !string.IsNullOrWhiteSpace(OutputPath)) {
                throw new PSArgumentException("Specify either -AsDocument or -OutputPath, not both.");
            }

            var html = Html;
            string? htmlFileDirectory = null;
            if (ParameterSetName == ParameterSetPath) {
                var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
                if (!File.Exists(resolvedPath)) {
                    throw new FileNotFoundException($"File '{resolvedPath}' was not found.", resolvedPath);
                }

                html = File.ReadAllText(resolvedPath);
                htmlFileDirectory = System.IO.Path.GetDirectoryName(resolvedPath);
            }

            if (string.IsNullOrWhiteSpace(html)) {
                ThrowTerminatingError(new ErrorRecord(
                    new ArgumentException("HTML content cannot be empty."),
                    "HtmlEmpty",
                    ErrorCategory.InvalidArgument,
                    html));
                return;
            }

            var options = BuildOptions(htmlFileDirectory);
            if (string.IsNullOrWhiteSpace(OutputPath) &&
                RequiresImageExtraction(options) &&
                !ShouldProcess(options.Base64ImageOutputDirectory, "Extract Markdown images from HTML")) {
                return;
            }

            if (AsDocument.IsPresent) {
                WriteObject(HtmlConversionDocument.Parse(html).ToMarkdownDocument(options));
                return;
            }

            if (!string.IsNullOrWhiteSpace(OutputPath)) {
                var resolvedOutput = SessionState.Path.GetUnresolvedProviderPathFromPSPath(OutputPath);
                if (!ShouldProcess(resolvedOutput, "Write Markdown converted from HTML")) {
                    return;
                }

                var markdown = HtmlConversionDocument.Parse(html).ToMarkdown(options);
                var directory = System.IO.Path.GetDirectoryName(resolvedOutput);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory)) {
                    Directory.CreateDirectory(directory);
                }

                File.WriteAllText(resolvedOutput, markdown, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
                if (PassThru.IsPresent) {
                    WriteObject(new FileInfo(resolvedOutput));
                }
            } else {
                var markdown = HtmlConversionDocument.Parse(html).ToMarkdown(options);
                WriteObject(markdown);
            }
        } catch (Exception ex) {
            WriteError(new ErrorRecord(ex, "HtmlToMarkdownFailed", ErrorCategory.InvalidOperation,
                ParameterSetName == ParameterSetPath ? Path : Html));
        }
    }

    private HtmlToMarkdownOptions BuildOptions(string? htmlFileDirectory) {
        var options = Options?.Clone()
            ?? (Portable.IsPresent
                ? HtmlToMarkdownOptions.CreatePortableProfile()
                : HtmlToMarkdownOptions.CreateOfficeIMOProfile());

        options.UseBodyContentsOnly = !IncludeDocumentChrome.IsPresent;
        options.RemoveScriptsAndStyles = !PreserveScriptsAndStyles.IsPresent;
        options.PreserveUnsupportedBlocks = !DropUnsupportedBlocks.IsPresent;
        options.PreserveUnsupportedInlineHtml = !DropUnsupportedInlineHtml.IsPresent;

        if (MaxInputCharacters.HasValue) {
            options.MaxInputCharacters = MaxInputCharacters.Value;
        }

        if (Base64ImageHandling.HasValue) {
            options.Base64Images = Base64ImageHandling.Value;
        }

        if (!string.IsNullOrWhiteSpace(Base64ImageOutputDirectory)) {
            options.Base64ImageOutputDirectory = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Base64ImageOutputDirectory);
        }

        if (ListingCardMetadataMode.HasValue) {
            options.ListingCardMetadataMode = ListingCardMetadataMode.Value;
        }

        if (MaxTableExpandedColumns.HasValue) {
            options.MaxTableExpandedColumns = MaxTableExpandedColumns.Value;
        }

        var writeOptions = MarkdownOptionUtilities.BuildWriteOptions(this);
        if (writeOptions != null) {
            options.MarkdownWriteOptions = writeOptions;
        }

        if (!string.IsNullOrWhiteSpace(BaseUri)) {
            options.BaseUri = new Uri(BaseUri, UriKind.Absolute);
        } else if (!string.IsNullOrWhiteSpace(htmlFileDirectory)) {
            options.BaseUri = new Uri(System.IO.Path.GetFullPath(htmlFileDirectory!) + System.IO.Path.DirectorySeparatorChar);
        }

        return options;
    }

    private static bool RequiresImageExtraction(HtmlToMarkdownOptions options)
        => options.Base64Images == HtmlBase64ImageHandling.SaveToFile &&
           !string.IsNullOrWhiteSpace(options.Base64ImageOutputDirectory);
}
