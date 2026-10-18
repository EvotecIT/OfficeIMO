using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Closes one or more tracked Word documents, optionally saving them.</summary>
/// <para>Provides a cmdlet wrapper for <c>WordDocument.Dispose</c>/<c>Save</c> so scripts need not call .NET methods directly.</para>
/// <example>
///   <summary>Close without saving.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$doc = Get-OfficeWord -Path .\Report.docx; Close-OfficeWord -Document $doc</code>
///   <para>Disposes the loaded document instance without saving changes.</para>
/// </example>
/// <example>
///   <summary>Close the most recently tracked document.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Close-OfficeWord</code>
///   <para>Closes the current tracked document when a document handle is not passed explicitly.</para>
/// </example>
/// <example>
///   <summary>Save to a new path and open the file.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Close-OfficeWord -Document $doc -Save -Path .\Report-final.docx -Open</code>
///   <para>Saves updates to <c>Report-final.docx</c>, opens it, and disposes the document.</para>
/// </example>
[Cmdlet(VerbsCommon.Close, "OfficeWord", DefaultParameterSetName = ParameterSetCurrent, SupportsShouldProcess = true)]
public sealed class CloseOfficeWordCommand : PSCmdlet {
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetCurrent = "Current";
    private const string ParameterSetAll = "All";

    /// <summary>Word document to close.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public WordDocument? Document { get; set; }

    /// <summary>Close the most recently tracked document.</summary>
    [Parameter(ParameterSetName = ParameterSetCurrent)]
    public SwitchParameter Current { get; set; }

    /// <summary>Close all tracked documents for the current runspace.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetAll)]
    public SwitchParameter All { get; set; }

    /// <summary>Persist changes before closing.</summary>
    [Parameter]
    public SwitchParameter Save { get; set; }

    /// <summary>Optional target path when saving.</summary>
    [Parameter(ParameterSetName = ParameterSetDocument)]
    [Parameter(ParameterSetName = ParameterSetCurrent)]
    public string? Path { get; set; }

    /// <summary>Open the file after saving. Requires -Save or -Path.</summary>
    [Parameter]
    [Alias("Show")]
    public SwitchParameter Open { get; set; }

    /// <summary>Password used to save the document as an encrypted package.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (Open.IsPresent && !Save.IsPresent && string.IsNullOrWhiteSpace(Path)) {
            throw new PSArgumentException("Use -Save or -Path with -Open so the document is persisted before it is opened.", nameof(Open));
        }

        if (All.IsPresent) {
            var documents = WordDocumentService.GetTrackedDocuments();
            for (var index = documents.Count - 1; index >= 0; index--) {
                CloseSingleDocument(documents[index]);
            }
            return;
        }

        WordDocument? document;
        if (ParameterSetName == ParameterSetDocument) {
            document = Document;
            if (document == null) {
                throw new PSArgumentNullException(nameof(Document), "Provide a WordDocument instance when using -Document.");
            }
        } else {
            document = WordDocumentService.GetCurrentTrackedDocument();
        }

        if (document == null) {
            throw new PSInvalidOperationException("No tracked Word document was found. Pass -Document or open a document with Get-OfficeWord/New-OfficeWord first.");
        }

        CloseSingleDocument(document);
    }

    private void CloseSingleDocument(WordDocument document) {
        var shouldSave = Save.IsPresent || !string.IsNullOrWhiteSpace(Path);
        var action = shouldSave ? "Save and close" : "Close";
        if (!ShouldProcess("Word document", action)) {
            return;
        }

        if (shouldSave) {
            var resolvedPath = !string.IsNullOrWhiteSpace(Path)
                ? SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path)
                : null;
            WordDocumentService.SaveDocument(document, Open.IsPresent, resolvedPath, Password);
            return;
        }

        WordDocumentService.CloseDocument(document);
    }
}
