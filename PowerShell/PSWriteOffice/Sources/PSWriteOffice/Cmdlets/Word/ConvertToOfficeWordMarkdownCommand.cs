using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Converts a Word document to Markdown.</summary>
/// <para>Returns Markdown text or saves it to a file when <c>-OutputPath</c> is specified.</para>
/// <example>
///   <summary>Convert a .docx file to Markdown text.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$markdown = ConvertTo-OfficeWordMarkdown -Path .\report.docx</code>
///   <para>Loads the document and returns Markdown markup.</para>
/// </example>
/// <example>
///   <summary>Save Markdown to disk.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ConvertTo-OfficeWordMarkdown -Path .\report.docx -OutputPath .\report.md -PassThru</code>
///   <para>Writes <c>report.md</c> and returns the file info.</para>
/// </example>
[Cmdlet(VerbsData.ConvertTo, "OfficeWordMarkdown", DefaultParameterSetName = ParameterSetPath, SupportsShouldProcess = true)]
[Alias("ConvertTo-WordMarkdown")]
[OutputType(typeof(string), typeof(FileInfo))]
public sealed class ConvertToOfficeWordMarkdownCommand : PSCmdlet {
    private const string ParameterSetPath = "Path";
    private const string ParameterSetDocument = "Document";

    /// <summary>Path to a .docx file.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Word document instance to convert.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public WordDocument Document { get; set; } = null!;

    /// <summary>Optional output path for the Markdown file.</summary>
    [Parameter]
    [Alias("OutPath")]
    public string? OutputPath { get; set; }

    /// <summary>Optional font family that should be treated as inline code.</summary>
    [Parameter]
    public string? FontFamily { get; set; }

    /// <summary>Wrap underlined text with HTML underline tags.</summary>
    [Parameter]
    public SwitchParameter EnableUnderline { get; set; }

    /// <summary>Wrap highlighted text with Markdown highlight markers.</summary>
    [Parameter]
    public SwitchParameter EnableHighlight { get; set; }

    /// <summary>Controls how images are emitted during Markdown conversion.</summary>
    [Parameter]
    public ImageExportMode ImageExportMode { get; set; } = ImageExportMode.Base64;

    /// <summary>Directory used when exporting images as files.</summary>
    [Parameter]
    public string? ImageDirectory { get; set; }

    /// <summary>Emit a <see cref="FileInfo"/> when saving to disk.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        WordDocument? document = null;
        var dispose = false;

        try {
            if (ParameterSetName == ParameterSetPath) {
                var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
                document = WordDocumentService.LoadDocument(resolvedPath, readOnly: true, autoSave: false);
                dispose = true;
            } else {
                document = Document;
            }

            if (document == null) {
                throw new InvalidOperationException("Word document was not provided.");
            }

            var options = new WordToMarkdownOptions {
                EnableUnderline = EnableUnderline.IsPresent,
                EnableHighlight = EnableHighlight.IsPresent,
                ImageExportMode = ImageExportMode
            };

            if (!string.IsNullOrWhiteSpace(FontFamily)) {
                options.FontFamily = FontFamily;
            }

            if (!string.IsNullOrWhiteSpace(ImageDirectory)) {
                options.ImageDirectory = SessionState.Path.GetUnresolvedProviderPathFromPSPath(ImageDirectory);
            }

            if (!string.IsNullOrWhiteSpace(OutputPath)) {
                var resolvedOutput = SessionState.Path.GetUnresolvedProviderPathFromPSPath(OutputPath);
                if (!ShouldProcess(resolvedOutput, "Write Markdown converted from Word document")) {
                    return;
                }

                var directory = System.IO.Path.GetDirectoryName(resolvedOutput);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory)) {
                    Directory.CreateDirectory(directory);
                }

                document.SaveAsMarkdown(resolvedOutput, options);
                if (PassThru.IsPresent) {
                    WriteObject(new FileInfo(resolvedOutput));
                }
            } else {
                if (options.ImageExportMode == ImageExportMode.File && !string.IsNullOrWhiteSpace(options.ImageDirectory) &&
                    !ShouldProcess(options.ImageDirectory, "Export Word images while converting to Markdown")) {
                    options.ImageExportMode = ImageExportMode.Base64;
                    options.ImageDirectory = null;
                }

                WriteObject(document.ToMarkdown(options));
            }
        } catch (Exception ex) {
            WriteError(new ErrorRecord(ex, "WordToMarkdownFailed", ErrorCategory.InvalidOperation,
                ParameterSetName == ParameterSetPath ? Path : Document));
        } finally {
            if (dispose) {
                document?.Dispose();
            }
        }
    }
}