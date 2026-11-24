using System.Management.Automation;
using OfficeIMO.Notebook;

namespace PSWriteOffice.Cmdlets.Notebook;

/// <summary>Clears saved results from a code cell, optionally including its execution count.</summary>
/// <example><summary>Clear all saved results before sharing a notebook.</summary><code>
/// $notebook | Get-OfficeNotebookCell -Kind Code | Clear-OfficeNotebookOutput -ExecutionCount
/// </code></example>
[Cmdlet(VerbsCommon.Clear, "OfficeNotebookOutput", SupportsShouldProcess = true)]
[OutputType(typeof(NotebookCell))]
public sealed class ClearOfficeNotebookOutputCommand : PSCmdlet {
    /// <summary>Code cell containing the saved results.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true), ValidateNotNull]
    public NotebookCell Cell { get; set; } = null!;

    /// <summary>Also clear the code cell's saved execution count.</summary>
    [Parameter]
    public SwitchParameter ExecutionCount { get; set; }

    /// <summary>Return the updated cell.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (!ShouldProcess(Cell.Id, "Clear saved notebook outputs")) return;
        Cell.ClearOutputs();
        if (ExecutionCount) Cell.ExecutionCount = null;
        if (PassThru) WriteObject(Cell);
    }
}
