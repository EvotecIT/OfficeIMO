using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Creates a Word document from HTML.</summary>
/// <para>Returns a <see cref="WordDocument"/> or saves it to disk when <c>-OutputPath</c> is provided.</para>
/// <example>
///   <summary>Create a .docx from HTML markup.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ConvertFrom-OfficeWordHtml -Html '&lt;h1&gt;Hello&lt;/h1&gt;' -OutputPath .\hello.docx</code>
///   <para>Writes a Word document containing the supplied HTML.</para>
/// </example>
/// <example>
///   <summary>Load HTML from disk and get the document object.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$doc = ConvertFrom-OfficeWordHtml -Path .\snippet.html</code>
///   <para>Returns a Word document instance for further edits.</para>
/// </example>
[Cmdlet(VerbsData.ConvertFrom, "OfficeWordHtml", DefaultParameterSetName = ParameterSetHtml, SupportsShouldProcess = true)]
[Alias("ConvertFrom-WordHtml")]
[OutputType(typeof(WordDocument), typeof(FileInfo))]
public sealed class ConvertFromOfficeWordHtmlCommand : PSCmdlet {
    private const string ParameterSetHtml = "Html";
    private const string ParameterSetPath = "Path";

    /// <summary>HTML markup to convert.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = ParameterSetHtml)]
    public string Html { get; set; } = string.Empty;

    /// <summary>Path to an HTML file.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Optional output path for the .docx file.</summary>
    [Parameter]
    [Alias("OutPath")]
    public string? OutputPath { get; set; }

    /// <summary>Optional font family to apply during conversion.</summary>
    [Parameter]
    public string? FontFamily { get; set; }

    /// <summary>Base path used to resolve relative resources (for example images).</summary>
    [Parameter]
    public string? BasePath { get; set; }

    /// <summary>Paths to CSS stylesheets to apply during conversion.</summary>
    [Parameter]
    public string[]? StylesheetPath { get; set; }

    /// <summary>Inline CSS stylesheets to apply during conversion.</summary>
    [Parameter]
    public string[]? StylesheetContent { get; set; }

    /// <summary>Include list style metadata.</summary>
    [Parameter]
    public SwitchParameter IncludeListStyles { get; set; }

    /// <summary>Continue numbering across separate ordered lists.</summary>
    [Parameter]
    public SwitchParameter ContinueNumbering { get; set; }

    /// <summary>Convert headings into a numbered list.</summary>
    [Parameter]
    public SwitchParameter SupportsHeadingNumbering { get; set; }

    /// <summary>Render &lt;pre&gt; elements as single-cell tables.</summary>
    [Parameter]
    public SwitchParameter RenderPreAsTable { get; set; }

    /// <summary>Controls where table captions are emitted.</summary>
    [Parameter]
    public TableCaptionPosition? TableCaptionPosition { get; set; }

    /// <summary>Controls how &lt;section&gt; tags are mapped into Word.</summary>
    [Parameter]
    public SectionTagHandling? SectionTagHandling { get; set; }

    /// <summary>Open the document after saving.</summary>
    [Parameter]
    public SwitchParameter Open { get; set; }

    /// <summary>Emit a <see cref="FileInfo"/> when saving to disk.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        try {
            var html = Html;
            string? htmlFileDirectory = null;
            if (ParameterSetName == ParameterSetPath) {
                var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
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

            var options = new HtmlToWordOptions {
                IncludeListStyles = IncludeListStyles.IsPresent,
                ContinueNumbering = ContinueNumbering.IsPresent,
                SupportsHeadingNumbering = SupportsHeadingNumbering.IsPresent,
                RenderPreAsTable = RenderPreAsTable.IsPresent
            };

            if (!string.IsNullOrWhiteSpace(FontFamily)) {
                options.FontFamily = FontFamily;
            }

            if (!string.IsNullOrWhiteSpace(BasePath)) {
                options.BasePath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(BasePath);
            } else if (!string.IsNullOrWhiteSpace(htmlFileDirectory)) {
                options.BasePath = htmlFileDirectory;
            }

            if (StylesheetPath != null) {
                foreach (var entry in StylesheetPath) {
                    if (string.IsNullOrWhiteSpace(entry)) {
                        continue;
                    }

                    options.StylesheetPaths.Add(SessionState.Path.GetUnresolvedProviderPathFromPSPath(entry));
                }
            }

            if (StylesheetContent != null) {
                foreach (var entry in StylesheetContent) {
                    if (!string.IsNullOrWhiteSpace(entry)) {
                        options.StylesheetContents.Add(entry);
                    }
                }
            }

            if (TableCaptionPosition.HasValue) {
                options.TableCaptionPosition = TableCaptionPosition.Value;
            }

            if (SectionTagHandling.HasValue) {
                options.SectionTagHandling = SectionTagHandling.Value;
            }

            var document = HtmlConversionDocument.Parse(html).ToWordDocument(options);

            if (!string.IsNullOrWhiteSpace(OutputPath)) {
                var resolvedOutput = SessionState.Path.GetUnresolvedProviderPathFromPSPath(OutputPath);
                if (!ShouldProcess(resolvedOutput, "Write Word document converted from HTML")) {
                    document.Dispose();
                    return;
                }

                var directory = System.IO.Path.GetDirectoryName(resolvedOutput);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory)) {
                    Directory.CreateDirectory(directory);
                }

                try {
                    document.Save(resolvedOutput);
                } finally {
                    document.Dispose();
                }

                if (Open.IsPresent) {
                    FileOpenService.Open(resolvedOutput);
                }

                if (PassThru.IsPresent) {
                    WriteObject(new FileInfo(resolvedOutput));
                }
            } else {
                WriteObject(document);
            }
        } catch (Exception ex) {
            WriteError(new ErrorRecord(ex, "HtmlToWordFailed", ErrorCategory.InvalidOperation,
                ParameterSetName == ParameterSetPath ? Path : Html));
        }
    }
}