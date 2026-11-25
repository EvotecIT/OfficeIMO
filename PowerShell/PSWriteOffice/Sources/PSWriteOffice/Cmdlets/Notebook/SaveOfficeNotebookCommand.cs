using System.Management.Automation;
using OfficeIMO.Core.Internal;
using OfficeIMO.Notebook;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Notebook;

/// <summary>Atomically saves a live notebook as .ipynb and leaves it open for further editing.</summary>
/// <example><summary>Save a new notebook or imported source to a new destination.</summary><code>
/// $notebook | Save-OfficeNotebook -Path ./updated.ipynb -NoClobber -PassThru
/// </code></example>
[Cmdlet(VerbsData.Save, "OfficeNotebook", SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo))]
public sealed class SaveOfficeNotebookCommand : AsyncPSCmdlet {
    /// <summary>Live notebook to save.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true), ValidateNotNull]
    public NotebookDocument Document { get; set; } = null!;

    /// <summary>.ipynb destination. Omission uses the path associated with a native notebook at creation or open time.</summary>
    [Parameter(Position = 0), ValidateNotNullOrEmpty]
    public string? Path { get; set; }

    /// <summary>Accept loss reported when normalizing the source to Jupyter notebook format.</summary>
    [Parameter]
    public SwitchParameter AllowLoss { get; set; }

    /// <summary>Refuse to replace a destination that already exists.</summary>
    [Parameter]
    public SwitchParameter NoClobber { get; set; }

    /// <summary>Return the saved file.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override async Task ProcessRecordAsync() {
        string destination = OfficePath.ResolveOutput(SessionState,
            Path ?? Document.FilePath ?? throw new PSArgumentException("Supply -Path when the notebook has no associated destination."), directory: false);
        NotebookCmdletPaths.RequireJupyterDestination(destination);
        if (!ShouldProcess(destination, "Save Jupyter notebook")) return;
        try {
            OfficePath.ValidateDestination(destination, NoClobber.IsPresent);
            await OfficeFileCommit.WriteAtomicallyAsync(destination,
                (stream, token) => Document.SaveAsync(stream, new NotebookSaveOptions { AllowLoss = AllowLoss.IsPresent }, token),
                CancelToken, NotebookCmdletPaths.ConflictPolicy(NoClobber.IsPresent)).ConfigureAwait(false);
            if (PassThru) WriteObject(new FileInfo(destination));
        } catch (Exception error) when (error is not (OperationCanceledException or PipelineStoppedException or ActionPreferenceStopException)) {
            WriteError(OfficeError.Create(error, "notebook-save-failed", destination));
        }
    }
}
