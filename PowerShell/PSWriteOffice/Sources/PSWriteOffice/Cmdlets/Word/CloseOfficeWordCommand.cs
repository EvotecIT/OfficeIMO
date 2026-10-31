using System.Management.Automation;
using System.Security;
using PSWriteOffice.Foundation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Closes one or more tracked Word documents, optionally saving them.</summary>
/// <para>Provides a cmdlet wrapper for <c>WordDocument.Dispose</c>/<c>Save</c> so scripts need not call .NET methods directly. Documents must use explicit persistence to prevent writes during disposal.</para>
/// <example>
///   <summary>Close without saving.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$doc = Open-OfficeWord -Path .\Report.docx; Close-OfficeWord -Document $doc</code>
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
///   <code>Close-OfficeWord -Document $doc -Save -Path .\Report-final.docx -Show</code>
///   <para>Saves updates to <c>Report-final.docx</c>, opens it, and disposes the document.</para>
/// </example>
[Cmdlet(VerbsCommon.Close, "OfficeWord", DefaultParameterSetName = ParameterSetCurrent, SupportsShouldProcess = true)]
public sealed class CloseOfficeWordCommand : OfficeCmdlet {
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetCurrent = "Current";
    private const string ParameterSetAll = "All";

    /// <summary>Word document to close.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public WordDocument? Document { get; set; }

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
    public SwitchParameter Show { get; set; }

    /// <summary>Explicitly accept reported native-save reductions.</summary>
    [Parameter]
    public SwitchParameter AllowLoss { get; set; }

    /// <summary>Password used to save the document as an encrypted package.</summary>
    [Parameter]
    [OfficePasswordTransformation]
    public SecureString? Password { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (Show.IsPresent && !Save.IsPresent && string.IsNullOrWhiteSpace(Path)) {
            throw new PSArgumentException("Use -Save or -Path with -Show so the document is persisted before it is opened.", nameof(Show));
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
            throw new PSInvalidOperationException("No tracked Word document was found. Pass -Document or open a document with Open-OfficeWord/New-OfficeWord first.");
        }

        CloseSingleDocument(document);
    }

    private void CloseSingleDocument(WordDocument document) {
        WordDocumentService.EnsureExplicitPersistence(document);
        var shouldSave = Save.IsPresent || !string.IsNullOrWhiteSpace(Path);
        if (shouldSave) {
            string? destination = Path ?? WordDocumentService.GetAssociatedPath(document);
            if (string.IsNullOrWhiteSpace(destination)) {
                InvalidArgument(new PSArgumentException("Supply Path when the document has no associated destination."), document);
                return;
            }
            destination = ResolveOutputPath(destination!);
            if (!ShouldProcess(destination, "Save and close Word document")) return;
            document.Save(destination, WordDocumentService.NativeSaveOptions(AllowLoss.IsPresent, Password,
                MyInvocation.BoundParameters.ContainsKey(nameof(Password))));
            WordDocumentService.CloseDocument(document);
            if (Show) OfficeShow.Open(destination);
            return;
        }

        WordDocumentService.CloseDocument(document);
    }
}
