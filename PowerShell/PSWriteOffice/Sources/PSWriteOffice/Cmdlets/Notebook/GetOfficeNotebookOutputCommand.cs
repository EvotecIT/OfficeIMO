using System.Management.Automation;
using OfficeIMO.Notebook;

namespace PSWriteOffice.Cmdlets.Notebook;

/// <summary>Returns saved notebook outputs in cell order without running any code.</summary>
/// <example><summary>Inspect the first output from the second cell.</summary><code>
/// $notebook | Get-OfficeNotebookCell -Index 1 | Get-OfficeNotebookOutput -Index 0
/// </code></example>
[Cmdlet(VerbsCommon.Get, "OfficeNotebookOutput")]
[OutputType(typeof(NotebookOutput))]
public sealed class GetOfficeNotebookOutputCommand : PSCmdlet {
    /// <summary>Cell containing the saved results.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true), ValidateNotNull]
    public NotebookCell Cell { get; set; } = null!;

    /// <summary>Zero-based saved output position. Omit to enumerate the available outputs.</summary>
    [Parameter, ValidateRange(0, int.MaxValue)]
    public int? Index { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (Index is int index) {
            WriteObject(Cell.Outputs[index]);
            return;
        }
        foreach (NotebookOutput output in Cell.Outputs) WriteObject(output);
    }
}
