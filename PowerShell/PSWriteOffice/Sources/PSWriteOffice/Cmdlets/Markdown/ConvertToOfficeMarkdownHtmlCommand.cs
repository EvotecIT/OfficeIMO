using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Drawing;
using OfficeIMO.Markdown;
using PSWriteOffice.Services.Markdown;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Converts Markdown content to HTML.</summary>
/// <para>Returns HTML text or saves it to a file when <c>-OutputPath</c> is specified.</para>
/// <example>
///   <summary>Convert a Markdown file to HTML.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$html = ConvertTo-OfficeMarkdownHtml -Path .\README.md</code>
///   <para>Returns the rendered HTML.</para>
/// </example>
/// <example>
///   <summary>Save a styled HTML document.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ConvertTo-OfficeMarkdownHtml -Path .\Report.md -DocumentMode -Title 'Weekly Report' -Style Clean -OutputPath .\Report.html -PassThru</code>
///   <para>Generates a full HTML file with title and CSS styling.</para>
/// </example>
[Cmdlet(VerbsData.ConvertTo, "OfficeMarkdownHtml", DefaultParameterSetName = ParameterSetPath, SupportsShouldProcess = true)]
[Alias("ConvertTo-MarkdownHtml")]
[OutputType(typeof(string), typeof(FileInfo))]
public sealed class ConvertToOfficeMarkdownHtmlCommand : PSCmdlet
    , IMarkdownReaderOptionSource {
    private const string ParameterSetPath = "Path";
    private const string ParameterSetText = "Text";
    private const string ParameterSetDocument = "Document";

    /// <summary>Path to the Markdown file.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Markdown text to convert.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetText)]
    public string Text { get; set; } = string.Empty;

    /// <summary>Markdown document to convert.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public MarkdownDoc Document { get; set; } = null!;

    /// <summary>Optional output path for the HTML file.</summary>
    [Parameter]
    [Alias("OutPath")]
    public string? OutputPath { get; set; }

    /// <summary>Render a full HTML document instead of a fragment.</summary>
    [Parameter]
    public SwitchParameter DocumentMode { get; set; }

    /// <summary>Built-in HTML style preset.</summary>
    [Parameter]
    public HtmlStyle Style { get; set; } = HtmlStyle.Clean;

    /// <summary>CSS delivery mode.</summary>
    [Parameter]
    public CssDelivery CssDelivery { get; set; } = CssDelivery.Inline;

    /// <summary>Asset loading mode.</summary>
    [Parameter]
    public AssetMode AssetMode { get; set; } = AssetMode.Online;

    /// <summary>Optional title for HTML documents.</summary>
    [Parameter]
    public string? Title { get; set; }

    /// <summary>Optional reader options when parsing Markdown.</summary>
    [Parameter]
    public MarkdownReaderOptions? ReaderOptions { get; set; }

    /// <summary>Named reader profile used when <see cref="ReaderOptions"/> is not supplied.</summary>
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

    /// <summary>Shared Markdown visual theme for HTML output.</summary>
    [Parameter]
    public OfficeVisualThemeKind? Theme { get; set; }

    /// <summary>Controls how raw HTML blocks are emitted.</summary>
    [Parameter]
    public RawHtmlHandling? RawHtmlHandling { get; set; }

    /// <summary>Add anchor links to headings.</summary>
    [Parameter]
    public SwitchParameter IncludeAnchorLinks { get; set; }

    /// <summary>Emit GitHub-compatible task-list HTML.</summary>
    [Parameter]
    public SwitchParameter GitHubTaskListHtml { get; set; }

    /// <summary>Emit GitHub-compatible footnote HTML.</summary>
    [Parameter]
    public SwitchParameter GitHubFootnoteHtml { get; set; }

    /// <summary>Open external HTTP(S) links in a new tab.</summary>
    [Parameter]
    public SwitchParameter ExternalLinksTargetBlank { get; set; }

    /// <summary>rel attribute value for external HTTP(S) links.</summary>
    [Parameter]
    public string? ExternalLinksRel { get; set; }

    /// <summary>referrerpolicy value for external HTTP(S) links.</summary>
    [Parameter]
    [ArgumentCompleter(typeof(OfficeReferrerPolicyArgumentCompleter))]
    public string? ExternalLinksReferrerPolicy { get; set; }

    /// <summary>Restrict absolute HTTP(S) links to the base origin.</summary>
    [Parameter]
    public SwitchParameter RestrictHttpLinksToBaseOrigin { get; set; }

    /// <summary>Restrict absolute HTTP(S) images to the base origin.</summary>
    [Parameter]
    public SwitchParameter RestrictHttpImagesToBaseOrigin { get; set; }

    /// <summary>Block all absolute external HTTP(S) images.</summary>
    [Parameter]
    public SwitchParameter BlockExternalHttpImages { get; set; }

    /// <summary>Add loading="lazy" to rendered images.</summary>
    [Parameter]
    public SwitchParameter ImagesLoadingLazy { get; set; }

    /// <summary>Add decoding="async" to rendered images.</summary>
    [Parameter]
    public SwitchParameter ImagesDecodingAsync { get; set; }

    /// <summary>referrerpolicy value for rendered images.</summary>
    [Parameter]
    [ArgumentCompleter(typeof(OfficeReferrerPolicyArgumentCompleter))]
    public string? ImagesReferrerPolicy { get; set; }

    /// <summary>Allowed HTTP(S) link hosts.</summary>
    [Parameter]
    public string[]? AllowedHttpLinkHost { get; set; }

    /// <summary>Allowed HTTP(S) image hosts.</summary>
    [Parameter]
    public string[]? AllowedHttpImageHost { get; set; }

    /// <summary>Emit a <see cref="FileInfo"/> when saving to disk.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var readerOptions = MarkdownOptionUtilities.BuildReaderOptions(this);

        MarkdownDoc document;
        if (ParameterSetName == ParameterSetDocument) {
            document = Document ?? throw new InvalidOperationException("Markdown document was not provided.");
        } else if (ParameterSetName == ParameterSetPath) {
            var resolved = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
            if (!File.Exists(resolved)) {
                throw new FileNotFoundException($"File '{resolved}' was not found.", resolved);
            }
            document = MarkdownDoc.Load(resolved, readerOptions);
        } else {
            document = MarkdownReader.Parse(Text ?? string.Empty, readerOptions);
        }

        var options = new HtmlOptions {
            Kind = DocumentMode.IsPresent ? HtmlKind.Document : HtmlKind.Fragment,
            Style = Style,
            CssDelivery = CssDelivery,
            AssetMode = AssetMode
        };

        ApplyHtmlOptions(options);

        if (!string.IsNullOrWhiteSpace(Title)) {
            options.Title = Title!;
        }

        if (!string.IsNullOrWhiteSpace(OutputPath)) {
            var resolvedOutput = SessionState.Path.GetUnresolvedProviderPathFromPSPath(OutputPath);
            if (!ShouldProcess(resolvedOutput, "Write HTML converted from Markdown")) {
                return;
            }

            var directory = System.IO.Path.GetDirectoryName(resolvedOutput);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory)) {
                Directory.CreateDirectory(directory);
            }

            document.SaveAsHtml(resolvedOutput, options);
            if (PassThru.IsPresent) {
                WriteObject(new FileInfo(resolvedOutput));
            }
        } else {
            WriteObject(options.Kind == HtmlKind.Document
                ? document.ToHtmlDocument(options)
                : document.ToHtmlFragment(options));
        }
    }

    private void ApplyHtmlOptions(HtmlOptions options) {
        if (Theme.HasValue) {
            options.Theme = MarkdownOptionUtilities.CreateTheme(Theme.Value);
        }

        if (RawHtmlHandling.HasValue) {
            options.RawHtmlHandling = RawHtmlHandling.Value;
        }

        options.IncludeAnchorLinks = IncludeAnchorLinks.IsPresent;
        options.GitHubTaskListHtml = GitHubTaskListHtml.IsPresent;
        options.GitHubFootnoteHtml = GitHubFootnoteHtml.IsPresent;
        options.ExternalLinksTargetBlank = ExternalLinksTargetBlank.IsPresent;
        options.RestrictHttpLinksToBaseOrigin = RestrictHttpLinksToBaseOrigin.IsPresent;
        options.RestrictHttpImagesToBaseOrigin = RestrictHttpImagesToBaseOrigin.IsPresent;
        options.BlockExternalHttpImages = BlockExternalHttpImages.IsPresent;
        options.ImagesLoadingLazy = ImagesLoadingLazy.IsPresent;
        options.ImagesDecodingAsync = ImagesDecodingAsync.IsPresent;

        if (!string.IsNullOrWhiteSpace(BaseUri) && Uri.TryCreate(BaseUri, UriKind.Absolute, out var baseUri)) {
            options.BaseUri = baseUri;
        }

        if (!string.IsNullOrWhiteSpace(ExternalLinksRel)) {
            options.ExternalLinksRel = ExternalLinksRel!;
        }

        if (!string.IsNullOrWhiteSpace(ExternalLinksReferrerPolicy)) {
            options.ExternalLinksReferrerPolicy = ExternalLinksReferrerPolicy!;
        }

        if (!string.IsNullOrWhiteSpace(ImagesReferrerPolicy)) {
            options.ImagesReferrerPolicy = ImagesReferrerPolicy!;
        }

        AddRange(options.AllowedHttpLinkHosts, AllowedHttpLinkHost);
        AddRange(options.AllowedHttpImageHosts, AllowedHttpImageHost);
    }

    private static void AddRange(System.Collections.Generic.ICollection<string> target, string[]? values) {
        if (values == null) {
            return;
        }

        foreach (var value in values) {
            if (!string.IsNullOrWhiteSpace(value)) {
                target.Add(value);
            }
        }
    }
}
