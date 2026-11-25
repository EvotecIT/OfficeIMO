using System.Management.Automation;
using System.Runtime.InteropServices;
using OfficeIMO.Core.Internal;
using OfficeIMO.Notebook;
using OfficeIMO.Notebook.PowerShell;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Notebook;

/// <summary>Runs selected PowerShell notebook cells in one persistent session and records their outputs.</summary>
/// <para>Run notebooks whose code you trust. Cells run under your account in a new runspace of the current PowerShell process, so they share its environment variables and loaded assemblies but not its variables or functions. Use -Isolation Process or -PowerShellPath to run them in a separate PowerShell process instead. The returned report contains the updated live notebook. Close the notebook when you finish using it.</para>
/// <para>Use -WhatIf to preview execution and any requested save. Use -Confirm to ask before running the notebook.</para>
/// <example><summary>Run PowerShell cells and save a separate notebook.</summary><code>
/// $report = Invoke-OfficeNotebook -LiteralPath ./lab.ipynb -OutputPath ./executed.ipynb -NoClobber
/// if ($null -ne $report) {
///     try {
///         $report.Cells
///     } finally {
///         $report.Notebook | Close-OfficeNotebook
///     }
/// }
/// </code></example>
/// <example><summary>Preview execution and ask for confirmation before running.</summary><code>
/// Invoke-OfficeNotebook -LiteralPath ./lab.ipynb -OutputPath ./reviewed.ipynb -WhatIf
/// $report = Invoke-OfficeNotebook -LiteralPath ./lab.ipynb -OutputPath ./reviewed.ipynb -Confirm
/// if ($null -ne $report) {
///     $report.Notebook | Close-OfficeNotebook
/// }
/// </code></example>
/// <example><summary>Run two selected cells in a live notebook.</summary><code>
/// $notebook | Invoke-OfficeNotebook -CellIndex 0, 2 -TimeoutSeconds 15
/// </code></example>
/// <example><summary>Run the cells in Windows PowerShell as a separate process.</summary><code>
/// $report = Invoke-OfficeNotebook -LiteralPath ./lab.ipynb -PowerShellPath "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
/// if ($null -ne $report) {
///     $report.Notebook | Close-OfficeNotebook
/// }
/// </code></example>
[Cmdlet(VerbsLifecycle.Invoke, "OfficeNotebook", DefaultParameterSetName = "Path", SupportsShouldProcess = true)]
[OutputType(typeof(NotebookExecutionReport), typeof(NotebookExecutionEvent))]
public sealed class InvokeOfficeNotebookCommand : AsyncPSCmdlet {
    private bool _outputWritten;

    /// <summary>Path to one native .ipynb notebook. FileSystem wildcards must resolve to a single file.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Path")]
    [SupportsWildcards, ValidateNotNullOrEmpty]
    public string Path { get; set; } = string.Empty;

    /// <summary>Literal .ipynb path, including a FileInfo.FullName or provider PSPath from the pipeline.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Alias("FullName", "PSPath"), ValidateNotNullOrEmpty]
    public string LiteralPath { get; set; } = string.Empty;

    /// <summary>Live notebook to update with the execution results.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "Document"), ValidateNotNull]
    public NotebookDocument Document { get; set; } = null!;

    /// <summary>Zero-based cell indexes. Selected PowerShell cells run in notebook order.</summary>
    [Parameter, ValidateNotNullOrEmpty, ValidateRange(0, int.MaxValue)]
    public int[]? CellIndex { get; set; }

    /// <summary>Maximum seconds per cell. The default is 60. Set to 0 to run without a per-cell timeout.</summary>
    [Parameter, ValidateRange(0, 2_147_483)]
    public int TimeoutSeconds { get; set; } = 60;

    /// <summary>Continue after PowerShell cell errors. A timeout or lost session stops execution.</summary>
    [Parameter]
    public SwitchParameter ContinueOnError { get; set; }

    /// <summary>Run cells without a declared cell or notebook language as PowerShell.</summary>
    [Parameter]
    public SwitchParameter IncludeUnspecifiedLanguage { get; set; }

    /// <summary>Where cells run. Runspace, the default, uses a new runspace in this PowerShell process. Process starts a separate PowerShell, which a timeout or Ctrl+C ends with any programs it started. Choose Process for cells that change process-wide state or can block in calls PowerShell cannot stop.</summary>
    [Parameter]
    public PowerShellNotebookIsolation Isolation { get; set; }

    /// <summary>Literal path to the PowerShell executable for a separate process. Implies -Isolation Process. The default uses the installed executable for this host when available.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? PowerShellPath { get; set; }

    /// <summary>Initial working directory. The default uses the notebook's directory when available.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? WorkingDirectory { get; set; }

    /// <summary>Write live output records to the pipeline before the final execution report.</summary>
    [Parameter]
    public SwitchParameter StreamOutput { get; set; }

    /// <summary>Atomically save the executed notebook to this .ipynb path, including any partial results reported after a cell error or timeout. Use this parameter with one notebook per invocation.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? OutputPath { get; set; }

    /// <summary>Refuse to replace an existing output file.</summary>
    [Parameter]
    public SwitchParameter NoClobber { get; set; }

    /// <summary>Accept loss reported when normalizing imported content for the .ipynb output.</summary>
    [Parameter]
    public SwitchParameter AllowLoss { get; set; }

    /// <inheritdoc />
    protected override async Task ProcessRecordAsync() {
        string? input = null;
        bool explicitIsolation = MyInvocation.BoundParameters.ContainsKey(nameof(Isolation));
        if (explicitIsolation && Isolation == PowerShellNotebookIsolation.Runspace && PowerShellPath != null)
            throw new PSArgumentException("-PowerShellPath starts a separate process. Remove it or use -Isolation Process.");
        if (ParameterSetName != "Document") {
            bool literal = ParameterSetName == "LiteralPath";
            IReadOnlyList<string> paths = OfficePath.ResolveInputs(this, new[] { literal ? LiteralPath : Path }, literal);
            if (paths.Count == 0) return;
            if (paths.Count != 1) throw new PSArgumentException("Select exactly one notebook per path.");
            input = paths[0];
            if (!string.Equals(System.IO.Path.GetExtension(input), ".ipynb", StringComparison.OrdinalIgnoreCase))
                throw new PSArgumentException("Use Open-OfficeNotebook to import non-.ipynb sources before running them.");
        }

        string? destination = OutputPath == null ? null : OfficePath.ResolveOutput(SessionState, OutputPath, directory: false);
        if (destination != null) {
            NotebookCmdletPaths.RequireJupyterDestination(destination);
            if (_outputWritten) throw new PSArgumentException("Use -OutputPath with one notebook per invocation.");
        } else if (NoClobber || AllowLoss) {
            throw new PSArgumentException("-NoClobber and -AllowLoss require -OutputPath.");
        }

        string notebookName = input ?? Document.FilePath ?? "live notebook";
        string target = destination == null ? notebookName : notebookName + " and save to " + destination;
        if (!ShouldProcess(target, "Execute PowerShell notebook cells")) return;

        NotebookDocument? owned = null;
        try {
            if (destination != null) OfficePath.ValidateDestination(destination, NoClobber.IsPresent);
            PowerShellNotebookSessionOptions sessionOptions = CaptureSessionOptions(input ?? Document.FilePath);
            if (input != null)
                owned = await NotebookDocument.LoadAsync(input, cancellationToken: CancelToken).ConfigureAwait(false);
            NotebookDocument notebook = owned ?? Document;
            var executionOptions = new NotebookExecutionOptions {
                CellIndices = CellIndex,
                CellTimeout = TimeoutSeconds == 0 ? null : TimeSpan.FromSeconds(TimeoutSeconds),
                ContinueOnError = ContinueOnError.IsPresent,
                IncludeUnspecifiedLanguage = IncludeUnspecifiedLanguage.IsPresent,
                OutputReceived = StreamOutput.IsPresent ? item => WriteObject(item) : null
            };
            using var session = await PowerShellNotebookSession.StartAsync(sessionOptions, CancelToken).ConfigureAwait(false);
            NotebookExecutionReport report = await session.ExecuteAsync(notebook, executionOptions, CancelToken).ConfigureAwait(false);
            if (destination != null) {
                await OfficeFileCommit.WriteAtomicallyAsync(destination,
                    (stream, token) => notebook.SaveAsync(stream, new NotebookSaveOptions { AllowLoss = AllowLoss.IsPresent }, token),
                    CancelToken, NotebookCmdletPaths.ConflictPolicy(NoClobber.IsPresent)).ConfigureAwait(false);
                _outputWritten = true;
            }
            WriteObject(report);
            owned = null;
        } catch (Exception error) when (error is not (OperationCanceledException or PipelineStoppedException or ActionPreferenceStopException)) {
            WriteError(OfficeError.Create(error, "notebook-execution-failed", notebookName));
        } finally {
            owned?.Dispose();
        }
    }

    private PowerShellNotebookSessionOptions CaptureSessionOptions(string? sourcePath) {
        PowerShellNotebookIsolation isolation = MyInvocation.BoundParameters.ContainsKey(nameof(Isolation)) ? Isolation
            : PowerShellPath == null ? PowerShellNotebookIsolation.Runspace : PowerShellNotebookIsolation.Process;
        string? executable = isolation == PowerShellNotebookIsolation.Runspace ? null
            : PowerShellPath == null ? GetHostExecutable() : OfficePath.ResolveInput(SessionState, PowerShellPath, literal: true)[0];
        string? notebookDirectory = sourcePath == null ? null : System.IO.Path.GetDirectoryName(sourcePath);
        string? directory = WorkingDirectory == null
            ? (Directory.Exists(notebookDirectory) ? notebookDirectory : SessionState.Path.CurrentFileSystemLocation.ProviderPath)
            : OfficePath.ResolveOutput(SessionState, WorkingDirectory, directory: true);
        return new PowerShellNotebookSessionOptions { Isolation = isolation, ExecutablePath = executable, WorkingDirectory = directory };
    }

    private string? GetHostExecutable() {
        if (SessionState.PSVariable.GetValue("PSHOME") is not string home) return null;
#if FRAMEWORK
        const string name = "powershell.exe";
#else
        string name = RuntimeInformation.IsOSPlatform(OSPlatform.Windows) ? "pwsh.exe" : "pwsh";
#endif
        string path = System.IO.Path.Combine(home, name);
        return File.Exists(path) ? path : null;
    }
}
