using System.Management.Automation;
using OfficeIMO.Notebook;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Notebook;

/// <summary>Creates a live Jupyter notebook that can be populated and saved without running cells.</summary>
/// <example><summary>Author a notebook and save it.</summary><code>
/// $notebook = New-OfficeNotebook -Path ./lab.ipynb
/// $notebook | Add-OfficeNotebookCell -Kind Markdown -Source '# Lab notes'
/// $notebook | Add-OfficeNotebookCell -Kind Code -Language PowerShell -Source 'Get-Date'
/// $notebook | Save-OfficeNotebook
/// $notebook | Close-OfficeNotebook
/// </code></example>
[Cmdlet(VerbsCommon.New, "OfficeNotebook")]
[OutputType(typeof(NotebookDocument))]
public sealed class NewOfficeNotebookCommand : OfficeCmdlet {
    /// <summary>Associates a default .ipynb save destination without writing it.</summary>
    [Parameter(Position = 0), ValidateNotNullOrEmpty]
    public string? Path { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        NotebookDocument notebook;
        if (Path == null) {
            notebook = NotebookDocument.Create();
        } else {
            string destination = ResolveOutputPath(Path);
            NotebookCmdletPaths.RequireJupyterDestination(destination);
            notebook = NotebookDocument.Create(destination);
        }
        try { WriteObject(notebook); }
        catch { notebook.Dispose(); throw; }
    }
}
