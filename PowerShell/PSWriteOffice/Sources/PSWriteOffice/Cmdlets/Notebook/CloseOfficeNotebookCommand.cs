using System.Management.Automation;
using OfficeIMO.Notebook;

namespace PSWriteOffice.Cmdlets.Notebook;

/// <summary>Closes a notebook and releases its source-backed values without writing a file.</summary>
/// <example><summary>Close a notebook after saving it.</summary><code>$notebook | Close-OfficeNotebook</code></example>
[Cmdlet(VerbsCommon.Close, "OfficeNotebook")]
public sealed class CloseOfficeNotebookCommand : PSCmdlet {
    /// <summary>Live notebook to close.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true), ValidateNotNull]
    public NotebookDocument Document { get; set; } = null!;

    /// <inheritdoc />
    protected override void ProcessRecord() => Document.Dispose();
}
