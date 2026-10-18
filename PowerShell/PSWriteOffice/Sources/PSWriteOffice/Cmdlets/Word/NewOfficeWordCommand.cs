using System;
using System.IO;
using System.Management.Automation;
using PSWriteOffice.Services.Pdf;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Creates a Word document using the DSL.</summary>
/// <para>Handles file creation or template cloning, scriptblock execution, explicit save or live-document composition, and emits the document path when <c>-PassThru</c> is used.</para>
/// <example>
///   <summary>Create a document inline.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeWord -Path .\Report.docx { WordSection { WordParagraph 'Hello DSL' } } -Open</code>
///   <para>Builds a document, adds one paragraph, saves it to disk, and opens it.</para>
/// </example>
/// <example>
///   <summary>Create a document from a template.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeWord -TemplatePath .\Template.docx -Path .\Report.docx { WordParagraph -Text 'Generated content' -StyleId 'ReportBody' }</code>
///   <para>Copies the template to the output path, runs the DSL against the copied document, and saves it.</para>
/// </example>
/// <example>
///   <summary>Keep a document for incremental composition.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$document = New-OfficeWord -Path .\Report.docx -NoSave
/// $document | Add-OfficeWordParagraph -Text 'Status report' -Style Heading1
/// $document | Save-OfficeWord
/// $document | Close-OfficeWord</code>
///   <para>Associates the output path with a live document, adds content through the pipeline, then saves and closes it once.</para>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeWord", SupportsShouldProcess = true)]
[Alias("OfficeWord", "WordNew")]
public sealed class NewOfficeWordCommand : PSCmdlet {
    /// <summary>Destination path for the document.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [Alias("FilePath", "OutputPath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Existing .docx file to clone before running the DSL.</summary>
    [Parameter]
    public string? TemplatePath { get; set; }

    /// <summary>DSL scriptblock describing document content.</summary>
    [Parameter(Position = 1)]
    public ScriptBlock? Content { get; set; }

    /// <summary>Emit a <see cref="FileInfo"/> for chaining.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <summary>Open the document after saving.</summary>
    [Parameter]
    public SwitchParameter Open { get; set; }

    /// <summary>Skip saving after executing the DSL.</summary>
    [Parameter]
    public SwitchParameter NoSave { get; set; }

    /// <summary>Password used to save the document as an encrypted package.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (NoSave.IsPresent && Open.IsPresent) {
            throw new PSArgumentException("-Open cannot be used with -NoSave because no file is written. Save the returned document explicitly, then use -Open on Save-OfficeWord.", nameof(Open));
        }

        var fullPath = GetResolvedPath();
        var action = NoSave.IsPresent
            ? string.IsNullOrWhiteSpace(TemplatePath)
                ? "Create in-memory Word document"
                : "Create Word document from template"
            : "Write new Word document";
        if (!PdfCommandUtilities.ShouldWrite(this, fullPath, action)) {
            return;
        }

        if (!NoSave.IsPresent || !string.IsNullOrWhiteSpace(TemplatePath)) {
            var directory = System.IO.Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory)) {
                Directory.CreateDirectory(directory);
            }
        }

        var document = CreateOrLoadDocument(fullPath);
        var closed = false;
        try {
            if (NoSave.IsPresent) {
                WordDocumentService.UpdateSaveAssociation(document, fullPath, encrypted: false);
            }

            if (Content != null) {
                WordDocumentService.InvokeDsl(document, Content);
            }

            if (NoSave.IsPresent) {
                WriteObject(document);
                return;
            }

            WordDocumentService.SaveDocument(document, Open.IsPresent, fullPath, Password);
            closed = true;

            if (PassThru.IsPresent) {
                WriteObject(new FileInfo(fullPath));
            }
        } catch {
            if (!closed) {
                WordDocumentService.CloseDocument(document);
            }
            throw;
        }
    }

    private string GetResolvedPath() {
        var providerPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
        return System.IO.Path.IsPathRooted(providerPath)
            ? providerPath
            : System.IO.Path.Combine(SessionState.Path.CurrentFileSystemLocation.Path, providerPath);
    }

    private OfficeIMO.Word.WordDocument CreateOrLoadDocument(string fullPath) {
        if (string.IsNullOrWhiteSpace(TemplatePath)) {
            if (NoSave.IsPresent) {
                return WordDocumentService.CreateInMemoryDocument();
            }

            return WordDocumentService.CreateDocument(fullPath, autoSave: false);
        }

        var templatePath = ResolveFileSystemPath(TemplatePath!);
        if (!File.Exists(templatePath)) {
            throw new FileNotFoundException($"Template file {templatePath} doesn't exist.", templatePath);
        }

        if (!string.Equals(templatePath, fullPath, StringComparison.OrdinalIgnoreCase)) {
            File.Copy(templatePath, fullPath, overwrite: true);
        }

        return WordDocumentService.LoadDocument(fullPath, readOnly: false, autoSave: false);
    }

    private string ResolveFileSystemPath(string path) {
        var providerPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(path);
        return System.IO.Path.IsPathRooted(providerPath)
            ? providerPath
            : System.IO.Path.Combine(SessionState.Path.CurrentFileSystemLocation.Path, providerPath);
    }

}
