using System.Management.Automation;
using OfficeIMO.Notebook;

namespace PSWriteOffice.Cmdlets.Notebook;

/// <summary>Appends a Markdown, code, or raw cell to a live notebook.</summary>
/// <example><summary>Append a PowerShell cell and return it for further editing.</summary><code>
/// $cell = $notebook | Add-OfficeNotebookCell -Kind Code -Language PowerShell -Source 'Get-Date' -PassThru
/// </code></example>
[Cmdlet(VerbsCommon.Add, "OfficeNotebookCell")]
[OutputType(typeof(NotebookCell))]
public sealed class AddOfficeNotebookCellCommand : PSCmdlet {
    /// <summary>Live notebook to extend.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true), ValidateNotNull]
    public NotebookDocument Document { get; set; } = null!;

    /// <summary>Kind of cell to append. Choose Markdown, Code, or Raw.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public NotebookCellKind Kind { get; set; }

    /// <summary>Cell text, which may be empty.</summary>
    [Parameter(Mandatory = true, Position = 1), AllowEmptyString]
    public string Source { get; set; } = string.Empty;

    /// <summary>Explicit language for a code cell.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? Language { get; set; }

    /// <summary>Return the appended cell.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (Language != null && Kind != NotebookCellKind.Code)
            throw new PSArgumentException("Language applies only to code cells.", nameof(Language));
        NotebookCell cell = Kind switch {
            NotebookCellKind.Markdown => Document.AddMarkdownCell(Source),
            NotebookCellKind.Code => Language == null ? Document.AddCodeCell(Source) : Document.AddCodeCell(Source, Language),
            NotebookCellKind.Raw => Document.AddRawCell(Source),
            _ => throw new PSArgumentException("Choose Markdown, Code, or Raw.", nameof(Kind))
        };
        if (PassThru) WriteObject(cell);
    }
}
