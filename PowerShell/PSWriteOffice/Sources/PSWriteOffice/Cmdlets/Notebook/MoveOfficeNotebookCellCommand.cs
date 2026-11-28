using System.Management.Automation;
using OfficeIMO.Notebook;

namespace PSWriteOffice.Cmdlets.Notebook;

/// <summary>Moves a cell within a live notebook while keeping its identity and saved content.</summary>
/// <example><summary>Move the third cell to the beginning.</summary><code>
/// $notebook | Move-OfficeNotebookCell -Index 2 -DestinationIndex 0
/// </code></example>
[Cmdlet(VerbsCommon.Move, "OfficeNotebookCell")]
[OutputType(typeof(NotebookDocument))]
public sealed class MoveOfficeNotebookCellCommand : PSCmdlet {
    /// <summary>Live notebook containing the cell.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true), ValidateNotNull]
    public NotebookDocument Document { get; set; } = null!;

    /// <summary>Original zero-based cell position.</summary>
    [Parameter(Mandatory = true, Position = 0), ValidateRange(0, int.MaxValue)]
    public int Index { get; set; }

    /// <summary>Zero-based final cell position.</summary>
    [Parameter(Mandatory = true, Position = 1), ValidateRange(0, int.MaxValue)]
    public int DestinationIndex { get; set; }

    /// <summary>Return the updated document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        Document.MoveCell(Index, DestinationIndex);
        if (PassThru) WriteObject(Document);
    }
}
