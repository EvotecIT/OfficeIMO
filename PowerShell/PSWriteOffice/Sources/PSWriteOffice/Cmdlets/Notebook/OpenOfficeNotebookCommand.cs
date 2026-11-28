using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Notebook;
using OfficeIMO.Notebook.Import;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Notebook;

/// <summary>Opens a Jupyter notebook or imports a saved notebook source into the editable model.</summary>
/// <para>Notebook cells remain inert. Close the returned document after using its source-backed values. Imported sources require an explicit .ipynb destination when saving.</para>
/// <example><summary>Open a native notebook and inspect code cells.</summary><code>
/// $notebook = Open-OfficeNotebook -LiteralPath ./lab.ipynb
/// $notebook | Get-OfficeNotebookCell -Kind Code
/// $notebook | Close-OfficeNotebook
/// </code></example>
/// <example><summary>Import a Jupytext script with a matching saved-output companion.</summary><code>
/// Open-OfficeNotebook -LiteralPath ./lab.py -Format JupytextPercent -CompanionPath ./lab.ipynb -FailOnLoss
/// </code></example>
[Cmdlet(VerbsCommon.Open, "OfficeNotebook", DefaultParameterSetName = "Path")]
[OutputType(typeof(NotebookDocument))]
public sealed class OpenOfficeNotebookCommand : AsyncPSCmdlet {
    /// <summary>Notebook path with FileSystem wildcard expansion.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Path")]
    [SupportsWildcards, ValidateNotNullOrEmpty]
    public string Path { get; set; } = string.Empty;

    /// <summary>Literal notebook path, including FileInfo.FullName or provider PSPath from a pipeline.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Alias("FullName", "PSPath"), ValidateNotNullOrEmpty]
    public string LiteralPath { get; set; } = string.Empty;

    /// <summary>Explicit import profile for ambiguous source extensions. Omission uses native .ipynb or an unambiguous import extension.</summary>
    [Parameter]
    public NotebookImportFormat? Format { get; set; }

    /// <summary>Literal path to a Jupyter notebook whose saved outputs may be matched to a Jupytext script.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? CompanionPath { get; set; }

    /// <summary>Exact notebook identifier within a Deepnote project.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? DeepnoteNotebookId { get; set; }

    /// <summary>Optional language override for Jupytext percent comment markers.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? ScriptLanguage { get; set; }

    /// <summary>Native resource limits or a hashtable of limit properties, applied to input and companion files.</summary>
    [Parameter, OfficeOptionsTransformation(typeof(NotebookLoadOptions)), ValidateNotNull]
    public NotebookLoadOptions? LoadOptions { get; set; }

    /// <summary>Reject source findings that describe fidelity loss.</summary>
    [Parameter]
    public SwitchParameter FailOnLoss { get; set; }

    /// <inheritdoc />
    protected override async Task ProcessRecordAsync() {
        bool literal = ParameterSetName == "LiteralPath";
        IReadOnlyList<string> paths = OfficePath.ResolveInputs(this, new[] { literal ? LiteralPath : Path }, literal);
        string? companion = CompanionPath == null ? null : OfficePath.ResolveInput(SessionState, CompanionPath, literal: true)[0];
        foreach (string path in paths) {
            NotebookDocument? notebook = null;
            try {
                NotebookImportFormat? profile = Format;
                if (profile == null && NotebookImporter.TryGetFormatFromPath(path, out var inferred)) profile = inferred;
                if (profile is { } selected) {
                    notebook = await ImportAsync(path, selected, companion).ConfigureAwait(false);
                } else {
                    if (!string.Equals(System.IO.Path.GetExtension(path), ".ipynb", StringComparison.OrdinalIgnoreCase))
                        throw new PSArgumentException("Specify -Format for this notebook source.");
                    if (companion != null || ScriptLanguage != null || DeepnoteNotebookId != null)
                        throw new PSArgumentException("CompanionPath, ScriptLanguage and DeepnoteNotebookId require their matching import profiles.");
                    notebook = await NotebookDocument.LoadAsync(path, LoadOptions, CancelToken).ConfigureAwait(false);
                }
                if (FailOnLoss)
                    new OfficeConversionReport(path, "Jupyter Notebook", diagnostics: notebook.Diagnostics).ThrowIfLoss();
                WriteObject(notebook);
                notebook = null;
            } catch (Exception error) when (error is not (OperationCanceledException or PipelineStoppedException or ActionPreferenceStopException)) {
                WriteError(OfficeError.Create(error, "notebook-open-failed", path));
            } finally {
                notebook?.Dispose();
            }
        }
    }

    private async Task<NotebookDocument> ImportAsync(string path, NotebookImportFormat profile, string? companion) {
        if (companion != null && profile != NotebookImportFormat.JupytextPercent)
            throw new PSArgumentException("CompanionPath requires the JupytextPercent profile.", nameof(CompanionPath));
        using NotebookDocument? paired = companion == null ? null :
            await NotebookDocument.LoadAsync(companion, LoadOptions, CancelToken).ConfigureAwait(false);
        var options = new NotebookImportOptions {
            Limits = LoadOptions ?? NotebookLoadOptions.UntrustedDefaults,
            PairedNotebook = paired,
            DeepnoteNotebookId = DeepnoteNotebookId,
            ScriptLanguage = ScriptLanguage
        };
        var imported = await NotebookImporter.LoadAsync(path, profile, options, CancelToken).ConfigureAwait(false);
        return imported.Value;
    }
}
