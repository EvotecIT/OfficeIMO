using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Pdf;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Saves a Word document without disposing it.</summary>
/// <para>Use <c>Close-OfficeWord -Save</c> when you want to save and dispose the document.</para>
/// <example>
///   <summary>Save the open document.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$doc | Save-OfficeWord</code>
///   <para>Persists pending changes and keeps the document open.</para>
/// </example>
[Cmdlet(VerbsData.Save, "OfficeWord", SupportsShouldProcess = true)]
[OutputType(typeof(WordDocument))]
public sealed class SaveOfficeWordCommand : PSCmdlet {
    /// <summary>Document to save.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
    public WordDocument Document { get; set; } = null!;

    /// <summary>Optional save-as path.</summary>
    [Parameter]
    [Alias("FilePath")]
    public string? Path { get; set; }

    /// <summary>Open the document after saving.</summary>
    [Parameter]
    [Alias("Show")]
    public SwitchParameter Open { get; set; }

    /// <summary>Password used to save the document as an encrypted package.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <summary>Emit the document object for further processing.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (Document == null) {
            return;
        }

        var associatedPath = WordDocumentService.GetAssociatedPath(Document);
        if (string.IsNullOrWhiteSpace(Path) && string.IsNullOrWhiteSpace(associatedPath)) {
            throw new PSInvalidOperationException("No file path provided. Use -Path or open the document from disk.");
        }

        string savedPath;
        if (!string.IsNullOrWhiteSpace(Path)) {
            var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
            if (!PdfCommandUtilities.ShouldWrite(this, resolvedPath, "Save Word document")) {
                return;
            }

            if (string.IsNullOrEmpty(Password) &&
                WordDocumentService.IsEncryptedSource(Document) &&
                string.Equals(System.IO.Path.GetFullPath(resolvedPath), System.IO.Path.GetFullPath(associatedPath!), StringComparison.OrdinalIgnoreCase)) {
                throw new PSInvalidOperationException("Provide -Password when saving a document loaded from an encrypted package.");
            }

            if (!string.IsNullOrEmpty(Password)) {
                OfficeEncryptedPackageService.SaveWord(Document, resolvedPath, Password!, false);
            } else {
                Document.Save(resolvedPath);
            }
            savedPath = resolvedPath;
        } else {
            if (!PdfCommandUtilities.ShouldWrite(this, associatedPath!, "Save Word document")) {
                return;
            }

            if (!string.IsNullOrEmpty(Password)) {
                OfficeEncryptedPackageService.SaveWord(Document, associatedPath!, Password!, false);
            } else {
                if (WordDocumentService.IsEncryptedSource(Document)) {
                    throw new PSInvalidOperationException("Provide -Password when saving a document loaded from an encrypted package.");
                }

                Document.Save(associatedPath!);
            }
            savedPath = associatedPath!;
        }

        WordDocumentService.UpdateSaveAssociation(Document, savedPath, !string.IsNullOrEmpty(Password));
        if (Open.IsPresent) {
            FileOpenService.Open(savedPath);
        }

        if (PassThru.IsPresent) {
            WriteObject(Document);
        }
    }

}