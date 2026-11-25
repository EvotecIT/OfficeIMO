using System.Management.Automation;
using OfficeIMO.Notebook;

namespace PSWriteOffice.Cmdlets.Notebook;

/// <summary>Streams notebook cells in document order, optionally by position or cell kind.</summary>
/// <example><summary>Select saved outputs from all code cells.</summary><code>
/// $notebook | Get-OfficeNotebookCell -Kind Code | Get-OfficeNotebookOutput
/// </code></example>
[Cmdlet(VerbsCommon.Get, "OfficeNotebookCell")]
[OutputType(typeof(NotebookCell))]
public sealed class GetOfficeNotebookCellCommand : PSCmdlet {
    /// <summary>Live notebook containing the cells.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true), ValidateNotNull]
    public NotebookDocument Document { get; set; } = null!;

    /// <summary>Zero-based cell position. Omit to enumerate the notebook.</summary>
    [Parameter, ValidateRange(0, int.MaxValue)]
    public int? Index { get; set; }

    /// <summary>Optional cell-kind filter.</summary>
    [Parameter]
    public NotebookCellKind? Kind { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (Index is int index) {
            NotebookCell cell = Document.Cells[index];
            if (Kind == null || cell.Kind == Kind) WriteObject(cell);
            return;
        }

        foreach (NotebookCell cell in Document.Cells) {
            if (Kind == null || cell.Kind == Kind) WriteObject(cell);
        }
    }
}
