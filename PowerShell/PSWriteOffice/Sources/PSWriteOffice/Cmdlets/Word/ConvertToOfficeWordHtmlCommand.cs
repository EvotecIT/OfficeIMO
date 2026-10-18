using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Converts a Word document to HTML.</summary>
/// <para>Returns the HTML string or saves it to a file when <c>-OutputPath</c> is specified.</para>
/// <example>
///   <summary>Convert a .docx file to HTML text.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$html = ConvertTo-OfficeWordHtml -Path .\report.docx</code>
///   <para>Loads the document and returns HTML markup.</para>
/// </example>
/// <example>
///   <summary>Save HTML to disk.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ConvertTo-OfficeWordHtml -Path .\report.docx -OutputPath .\report.html -PassThru</code>
///   <para>Writes <c>report.html</c> and returns the file info.</para>
/// </example>
[Cmdlet(VerbsData.ConvertTo, "OfficeWordHtml", DefaultParameterSetName = ParameterSetPath, SupportsShouldProcess = true)]
[Alias("ConvertTo-WordHtml")]
[OutputType(typeof(string), typeof(FileInfo))]
public sealed class ConvertToOfficeWordHtmlCommand : PSCmdlet {
    private const string ParameterSetPath = "Path";
    private const string ParameterSetDocument = "Document";

    /// <summary>Path to a .docx file.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Word document instance to convert.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public WordDocument Document { get; set; } = null!;

    /// <summary>Optional output path for the HTML file.</summary>
    [Parameter]
    [Alias("OutPath")]
    public string? OutputPath { get; set; }

    /// <summary>Optional font family to use during conversion.</summary>
    [Parameter]
    public string? FontFamily { get; set; }

    /// <summary>Include font styles as inline CSS.</summary>
    [Parameter]
    public SwitchParameter IncludeFontStyles { get; set; }

    /// <summary>Include list style metadata.</summary>
    [Parameter]
    public SwitchParameter IncludeListStyles { get; set; }

    /// <summary>Emit paragraph styles as CSS classes.</summary>
    [Parameter]
    public SwitchParameter IncludeParagraphClasses { get; set; }

    /// <summary>Emit run styles as CSS classes.</summary>
    [Parameter]
    public SwitchParameter IncludeRunClasses { get; set; }

    /// <summary>Include the built-in default CSS in the HTML head.</summary>
    [Parameter]
    public SwitchParameter IncludeDefaultCss { get; set; }

    /// <summary>Store image references as file paths instead of base64 data URIs.</summary>
    [Parameter]
    public SwitchParameter UseImagePaths { get; set; }

    /// <summary>Exclude footnotes from the HTML output.</summary>
    [Parameter]
    public SwitchParameter ExcludeFootnotes { get; set; }

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

            var options = new WordToHtmlOptions {
                IncludeFontStyles = IncludeFontStyles.IsPresent,
                IncludeListStyles = IncludeListStyles.IsPresent,
                IncludeParagraphClasses = IncludeParagraphClasses.IsPresent,
                IncludeRunClasses = IncludeRunClasses.IsPresent,
                IncludeDefaultCss = IncludeDefaultCss.IsPresent
            };

            if (!string.IsNullOrWhiteSpace(FontFamily)) {
                options.FontFamily = FontFamily;
            }

            if (UseImagePaths.IsPresent) {
                options.EmbedImagesAsBase64 = false;
            }

            if (ExcludeFootnotes.IsPresent) {
                options.ExportFootnotes = false;
            }

            if (!string.IsNullOrWhiteSpace(OutputPath)) {
                var resolvedOutput = SessionState.Path.GetUnresolvedProviderPathFromPSPath(OutputPath);
                if (!ShouldProcess(resolvedOutput, "Write HTML converted from Word document")) {
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
                WriteObject(document.ToHtml(options));
            }
        } catch (Exception ex) {
            WriteError(new ErrorRecord(ex, "WordToHtmlFailed", ErrorCategory.InvalidOperation,
                ParameterSetName == ParameterSetPath ? Path : Document));
        } finally {
            if (dispose) {
                document?.Dispose();
            }
        }
    }
}