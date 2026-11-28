using System.Management.Automation;
using OfficeIMO.Notebook;

namespace PSWriteOffice.Cmdlets.Notebook;

/// <summary>Attaches saved output or supplied console text to a code cell.</summary>
/// <example><summary>Record console text captured by the caller.</summary><code>
/// $cell | Add-OfficeNotebookOutput -Stream stdout -Text 'Lab result'
/// </code></example>
[Cmdlet(VerbsCommon.Add, "OfficeNotebookOutput", DefaultParameterSetName = "Stream")]
[OutputType(typeof(NotebookCell))]
public sealed class AddOfficeNotebookOutputCommand : PSCmdlet {
    /// <summary>Code cell to extend.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true), ValidateNotNull]
    public NotebookCell Cell { get; set; } = null!;

    /// <summary>Existing saved output to copy to the cell.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "Saved"), ValidateNotNull]
    public NotebookOutput Output { get; set; } = null!;

    /// <summary>Console stream associated with the supplied text.</summary>
    [Parameter(ParameterSetName = "Stream"), ValidateSet("stdout", "stderr")]
    public string Stream { get; set; } = "stdout";

    /// <summary>Saved text for the stream output.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "Stream"), AllowEmptyString]
    public string Text { get; set; } = string.Empty;

    /// <summary>Return the updated cell.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var output = ParameterSetName == "Saved" ? Output : NotebookOutput.CreateStream(Stream.ToLowerInvariant(), Text);
        Cell.AddOutput(output);
        if (PassThru) WriteObject(Cell);
    }
}
