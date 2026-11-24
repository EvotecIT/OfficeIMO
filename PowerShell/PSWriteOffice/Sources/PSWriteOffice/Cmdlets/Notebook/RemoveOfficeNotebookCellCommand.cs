using System.Management.Automation;
using OfficeIMO.Notebook;

namespace PSWriteOffice.Cmdlets.Notebook;

/// <summary>Removes a cell from a live notebook.</summary>
/// <example><summary>Preview removal of the first cell.</summary><code>
/// $notebook | Remove-OfficeNotebookCell -Index 0 -WhatIf
/// </code></example>
[Cmdlet(VerbsCommon.Remove, "OfficeNotebookCell", SupportsShouldProcess = true)]
[OutputType(typeof(NotebookDocument))]
public sealed class RemoveOfficeNotebookCellCommand : PSCmdlet {
    /// <summary>Live notebook containing the cell.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true), ValidateNotNull]
    public NotebookDocument Document { get; set; } = null!;

    /// <summary>Zero-based position of the cell to remove.</summary>
    [Parameter(Mandatory = true, Position = 0), ValidateRange(0, int.MaxValue)]
    public int Index { get; set; }

    /// <summary>Return the updated document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        NotebookCell cell = Document.Cells[Index];
        if (!ShouldProcess(cell.Id, "Remove notebook cell")) return;
        Document.RemoveCell(Index);
        if (PassThru) WriteObject(Document);
    }
}
