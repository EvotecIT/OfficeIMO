using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Core.Internal;
using OfficeIMO.Notebook;
using OfficeIMO.Notebook.Html;
using OfficeIMO.Notebook.Import;
using OfficeIMO.Notebook.Markdown;
using OfficeIMO.Notebook.Pdf;
using OfficeIMO.Notebook.Spreadsheet;
using OfficeIMO.Notebook.Word;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Notebook;

/// <summary>Exports a live notebook or a selected saved structured output, returning the conversion report.</summary>
/// <para>Use .pdf, .html, .md, .docx or .ps1 for live notebooks. Pipe a saved structured table to write .csv or .xlsx. The command never runs cells and preserves an existing destination if conversion fails.</para>
/// <example><summary>Export a PDF without showing code.</summary><code>
/// $notebook | Export-OfficeNotebook -Path ./lab.pdf -IncludeCode:$false -Title 'Lab report'
/// </code></example>
/// <example><summary>Export the first saved structured output from one code cell.</summary><code>
/// $notebook | Get-OfficeNotebookCell -Index 2 | Get-OfficeNotebookOutput -Index 0 | Export-OfficeNotebook -Path ./results.xlsx
/// </code></example>
/// <example><summary>Extract PowerShell source without running the notebook.</summary><code>
/// $notebook | Export-OfficeNotebook -Path ./lab.ps1 -FailOnLoss
/// </code></example>
[Cmdlet(VerbsData.Export, "OfficeNotebook", DefaultParameterSetName = "Document", SupportsShouldProcess = true)]
[OutputType(typeof(OfficeConversionReport))]
public sealed class ExportOfficeNotebookCommand : AsyncPSCmdlet {
    /// <summary>Live notebook to export to a document or a PowerShell script.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "Document"), ValidateNotNull]
    public NotebookDocument Document { get; set; } = null!;

    /// <summary>Selected saved output containing an inline structured data resource.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "SavedTable"), ValidateNotNull]
    public NotebookOutput Output { get; set; } = null!;

    /// <summary>Output path. The extension selects the format.</summary>
    [Parameter(Mandatory = true, Position = 0), ValidateNotNullOrEmpty]
    public string Path { get; set; } = string.Empty;

    /// <summary>Include code in a rendered document. Enabled by default. Use -IncludeCode:$false to omit it.</summary>
    [Parameter(ParameterSetName = "Document")]
    public SwitchParameter IncludeCode { get; set; }

    /// <summary>Include saved outputs in a rendered document. Enabled by default.</summary>
    [Parameter(ParameterSetName = "Document")]
    public SwitchParameter IncludeOutputs { get; set; }

    /// <summary>Include saved execution counts in a rendered document. Enabled by default.</summary>
    [Parameter(ParameterSetName = "Document")]
    public SwitchParameter IncludeExecutionCounts { get; set; }

    /// <summary>Preserve exact Markdown cell source in a Markdown export. Cell-local attachments are omitted with a report.</summary>
    [Parameter(ParameterSetName = "Document")]
    public SwitchParameter PreserveMarkdownSource { get; set; }

    /// <summary>Include supported, interactive saved-state widgets in a standalone HTML export.</summary>
    [Parameter(ParameterSetName = "Document")]
    public SwitchParameter IncludeInteractiveWidgets { get; set; }

    /// <summary>Enable zoom, pan and saved data inspection in supported offline HTML charts.</summary>
    [Parameter(ParameterSetName = "Document")]
    public SwitchParameter IncludeInteractiveCharts { get; set; }

    /// <summary>Optional title for a rendered document.</summary>
    [Parameter(ParameterSetName = "Document")]
    public string? Title { get; set; }

    /// <summary>Optional author for a rendered document.</summary>
    [Parameter(ParameterSetName = "Document")]
    public string? Author { get; set; }

    /// <summary>Reject a notebook export that reports fidelity loss before publishing a file.</summary>
    [Parameter(ParameterSetName = "Document")]
    public SwitchParameter FailOnLoss { get; set; }

    /// <summary>Refuse to replace an existing destination.</summary>
    [Parameter]
    public SwitchParameter NoClobber { get; set; }

    /// <inheritdoc />
    protected override async Task ProcessRecordAsync() {
        string destination = OfficePath.ResolveOutput(SessionState, Path, directory: false);
        string extension = System.IO.Path.GetExtension(destination).ToLowerInvariant();
        var write = ParameterSetName == "SavedTable" ? CaptureTableExport(extension) : CaptureDocumentExport(extension);
        if (!ShouldProcess(destination, "Export notebook content")) return;

        try {
            OfficePath.ValidateDestination(destination, NoClobber.IsPresent);
            OfficeConversionReport? report = null;
            await OfficeFileCommit.WriteAtomicallyAsync(destination, async (stream, token) => {
                report = await write(stream, token).ConfigureAwait(false);
            }, CancelToken, NotebookCmdletPaths.ConflictPolicy(NoClobber.IsPresent)).ConfigureAwait(false);
            WriteObject(report!.WithOutputPath(destination));
        } catch (Exception error) when (error is not (OperationCanceledException or PipelineStoppedException or ActionPreferenceStopException)) {
            WriteError(OfficeError.Create(error, "notebook-export-failed", destination));
        }
    }

    private Func<Stream, CancellationToken, Task<OfficeConversionReport>> CaptureTableExport(string extension) {
        NotebookOutput output = Output;
        return extension switch {
            ".csv" => (stream, token) => Task.FromResult(output.SaveAsCsv(stream, cancellationToken: token)),
            ".xlsx" => (stream, token) => Task.FromResult(output.SaveAsExcel(stream, token)),
            _ => throw new PSArgumentException("Export a saved structured output to .csv or .xlsx.")
        };
    }

    private Func<Stream, CancellationToken, Task<OfficeConversionReport>> CaptureDocumentExport(string extension) {
        NotebookDocument document = Document;
        var parameters = MyInvocation.BoundParameters;
        if (parameters.ContainsKey(nameof(PreserveMarkdownSource)) && extension is not (".md" or ".markdown")) {
            throw new PSArgumentException("-PreserveMarkdownSource applies only to Markdown exports.");
        }
        if (extension is not (".html" or ".htm")) {
            if (parameters.ContainsKey(nameof(IncludeInteractiveWidgets))) {
                throw new PSArgumentException("-IncludeInteractiveWidgets applies only to HTML exports.");
            }
            if (parameters.ContainsKey(nameof(IncludeInteractiveCharts))) {
                throw new PSArgumentException("-IncludeInteractiveCharts applies only to HTML exports.");
            }
        }
        bool failOnLoss = FailOnLoss.IsPresent;
        bool includeInteractiveWidgets = IncludeInteractiveWidgets.IsPresent;
        bool includeInteractiveCharts = IncludeInteractiveCharts.IsPresent;
        if (extension == ".ps1") {
            if (parameters.ContainsKey(nameof(IncludeCode)) || parameters.ContainsKey(nameof(IncludeOutputs)) ||
                parameters.ContainsKey(nameof(IncludeExecutionCounts)) || parameters.ContainsKey(nameof(Title)) ||
                parameters.ContainsKey(nameof(Author))) {
                throw new PSArgumentException("Document presentation options do not apply to PowerShell scripts.");
            }
            return (stream, token) => NotebookPowerShellScript.SaveAsync(document, stream, failOnLoss, token);
        }

        var projection = new NotebookMarkdownExportOptions {
            IncludeCode = OfficeSwitch.Get(parameters, nameof(IncludeCode)) ?? true,
            IncludeOutputs = OfficeSwitch.Get(parameters, nameof(IncludeOutputs)) ?? true,
            IncludeExecutionCounts = OfficeSwitch.Get(parameters, nameof(IncludeExecutionCounts)) ?? true,
            PreserveMarkdownSource = PreserveMarkdownSource.IsPresent,
            Title = Title, Author = Author, FailOnLoss = failOnLoss
        };
        return extension switch {
            ".md" or ".markdown" => (stream, token) => document.SaveAsMarkdownAsync(stream, projection, token),
            ".html" or ".htm" => (stream, token) => document.SaveAsHtmlAsync(stream,
                new NotebookHtmlExportOptions {
                    ProjectionOptions = projection,
                    FailOnLoss = failOnLoss,
                    IncludeInteractiveWidgets = includeInteractiveWidgets,
                    IncludeInteractiveCharts = includeInteractiveCharts
                }, token),
            ".pdf" => (stream, token) => document.SaveAsPdfAsync(stream, new NotebookPdfExportOptions {
                IncludeCode = projection.IncludeCode, IncludeOutputs = projection.IncludeOutputs,
                IncludeExecutionCounts = projection.IncludeExecutionCounts, Title = projection.Title,
                Author = projection.Author, FailOnLoss = failOnLoss
            }, token),
            ".docx" => (stream, token) => document.SaveAsWordAsync(stream, new NotebookWordExportOptions {
                IncludeCode = projection.IncludeCode, IncludeOutputs = projection.IncludeOutputs,
                IncludeExecutionCounts = projection.IncludeExecutionCounts, Title = projection.Title,
                Author = projection.Author, FailOnLoss = failOnLoss
            }, token),
            _ => throw new PSArgumentException("Export a notebook to .pdf, .html, .md, .docx, or .ps1.")
        };
    }
}
