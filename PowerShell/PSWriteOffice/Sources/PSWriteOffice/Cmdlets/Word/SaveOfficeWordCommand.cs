using System.IO;
using System.Management.Automation;
using System.Security;
using OfficeIMO.Word;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Saves a live Word document without closing it.</summary>
/// <para>Uses the supplied path or associated destination. Omitted Password preserves loaded encryption. PassThru returns the saved FileInfo. The document must use explicit persistence.</para>
/// <example><summary>Save and retain the live document for further editing.</summary>
/// <code>Save-OfficeWord -Document $document -Path ./Report.docx -PassThru
/// Close-OfficeWord -Document $document</code></example>
/// <example><summary>Save a loaded document with its existing encryption.</summary>
/// <code>$document = Open-OfficeWord -Path ./Protected.docx -Password (Read-Host -AsSecureString)
/// $document | Save-OfficeWord
/// $document | Close-OfficeWord</code></example>
[Cmdlet(VerbsData.Save, "OfficeWord", SupportsShouldProcess = true, DefaultParameterSetName = "Document")]
[OutputType(typeof(FileInfo))]
public sealed class SaveOfficeWordCommand : OfficeCmdlet {
    /// <summary>Live Word document to save.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
    [ValidateNotNull]
    public WordDocument Document { get; set; } = null!;

    /// <summary>Optional new destination; omission uses the associated path.</summary>
    [Parameter]
    [ValidateNotNullOrEmpty]
    public string? Path { get; set; }

    /// <summary>Override the output password; omission preserves the document's loaded encryption.</summary>
    [Parameter]
    [OfficePasswordTransformation]
    public SecureString? Password { get; set; }

    /// <summary>Explicitly accept reported native-save reductions.</summary>
    [Parameter]
    public SwitchParameter AllowLoss { get; set; }

    /// <summary>Open the saved file in its registered default application.</summary>
    [Parameter]
    public SwitchParameter Show { get; set; }

    /// <summary>Return the saved FileInfo.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        WordDocumentService.EnsureExplicitPersistence(Document);
        string? destination = Path ?? WordDocumentService.GetAssociatedPath(Document);
        if (string.IsNullOrWhiteSpace(destination)) {
            InvalidArgument(new PSArgumentException("Supply Path when the document has no associated destination."), Document);
            return;
        }
        destination = ResolveOutputPath(destination!);
        if (!ShouldProcess(destination, "Save Word document")) return;
        try {
            bool passwordBound = MyInvocation.BoundParameters.ContainsKey(nameof(Password));
            Document.Save(destination, WordDocumentService.NativeSaveOptions(AllowLoss.IsPresent, Password, passwordBound));
            bool encrypted = passwordBound ? Password != null : WordDocumentService.IsEncryptedSource(Document);
            WordDocumentService.UpdateSaveAssociation(Document, destination, encrypted);
            if (Show) OfficeShow.Open(destination);
            if (PassThru) WriteObject(new FileInfo(destination));
        } catch (PipelineStoppedException) { throw; }
        catch (System.Exception error) { ItemError(error, "native-save-failed", destination); }
    }
}
