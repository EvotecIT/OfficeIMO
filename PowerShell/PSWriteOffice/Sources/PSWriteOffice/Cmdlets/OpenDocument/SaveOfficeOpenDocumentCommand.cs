using System.IO;
using System.Management.Automation;
using OfficeIMO.OpenDocument;

namespace PSWriteOffice.Cmdlets.OpenDocument;

/// <summary>Saves a native OpenDocument model with entry-level preservation diagnostics.</summary>
/// <example><summary>Save a live document to its associated destination.</summary>
/// <code>$document = New-OfficeOpenDocument -Path ./Proposal.odt
/// Save-OfficeOpenDocument -Document $document</code></example>
[Cmdlet(VerbsData.Save, "OfficeOpenDocument", SupportsShouldProcess = true)]
[OutputType(typeof(OdfSaveResult))]
public sealed class SaveOfficeOpenDocumentCommand : PSCmdlet {
    /// <summary>OpenDocument model to save.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public OdfDocument Document { get; set; } = null!;

    /// <summary>Destination path; defaults to the document's associated save path.</summary>
    [Parameter(Position = 0), ValidateNotNullOrEmpty]
    public string? Path { get; set; }

    /// <summary>Optional package save and preservation settings.</summary>
    [Parameter]
    public OdfSaveOptions? Options { get; set; }

    /// <summary>Throw when source entries cannot be preserved losslessly.</summary>
    [Parameter]
    public SwitchParameter FailOnLoss { get; set; }

    /// <summary>Emit the save result, including preservation diagnostics.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        string destination = Path ?? Document.FilePath ?? throw new PSArgumentException("Supply Path when the document has no associated save path.", nameof(Path));
        var output = SessionState.Path.GetUnresolvedProviderPathFromPSPath(destination);
        OpenDocumentCommandUtilities.ValidateOpenDocumentExtension(output, Document.Kind, nameof(Path));
        if (!ShouldProcess(output, "Save OpenDocument package")) return;
        if (FailOnLoss.IsPresent) {
            Document.Serialize(Options).RequireNoLoss();
        }
        Directory.CreateDirectory(System.IO.Path.GetDirectoryName(output) ?? SessionState.Path.CurrentFileSystemLocation.Path);
        var result = Document.Save(output, Options);
        if (PassThru.IsPresent) {
            WriteObject(result);
        }
    }
}
