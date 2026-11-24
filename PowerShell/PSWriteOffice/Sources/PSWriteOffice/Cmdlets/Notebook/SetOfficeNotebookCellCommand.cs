using System.Management.Automation;
using OfficeIMO.Notebook;

namespace PSWriteOffice.Cmdlets.Notebook;

/// <summary>Updates cell text or a saved code-cell execution count without running code.</summary>
/// <example><summary>Replace one cell's source and clear its execution count.</summary><code>
/// $notebook | Get-OfficeNotebookCell -Index 1 | Set-OfficeNotebookCell -Source 'Get-Date' -ExecutionCount $null
/// </code></example>
[Cmdlet(VerbsCommon.Set, "OfficeNotebookCell")]
[OutputType(typeof(NotebookCell))]
public sealed class SetOfficeNotebookCellCommand : PSCmdlet {
    /// <summary>Cell to edit. Keep its notebook open until the edit is complete.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true), ValidateNotNull]
    public NotebookCell Cell { get; set; } = null!;

    /// <summary>Replacement cell text.</summary>
    [Parameter, AllowEmptyString, ValidateNotNull]
    public string? Source { get; set; }

    /// <summary>Saved execution count for a code cell. Supply $null to clear it.</summary>
    [Parameter, AllowNull]
    public int? ExecutionCount { get; set; }

    /// <summary>Return the updated cell.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        bool sourceBound = MyInvocation.BoundParameters.ContainsKey(nameof(Source));
        bool countBound = MyInvocation.BoundParameters.ContainsKey(nameof(ExecutionCount));
        if (!sourceBound && !countBound)
            throw new PSArgumentException("Specify -Source or -ExecutionCount.");
        if (countBound) Cell.ExecutionCount = ExecutionCount;
        if (sourceBound) Cell.Source = Source!;
        if (PassThru) WriteObject(Cell);
    }
}
