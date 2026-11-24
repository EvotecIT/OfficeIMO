using OfficeIMO.Notebook.PowerShell;

namespace OfficeIMO.Notebook.Tests;

public sealed class NotebookPowerShellExecutionTests {
    [PowerShellRuntimeFact]
    public async Task SelectedCellsRetainStateAcrossRunsAndSaveOnlyAttemptedOutputs() {
        using var notebook = NotebookDocument.Create();
        notebook.SetMetadata("kernelspec", JsonSerializer.SerializeToElement(new { name = "powershell" }));
        notebook.AddMarkdownCell("# Lab");
        NotebookCell setup = notebook.AddCodeCell("$n = 41; function Next-Number { (++$script:n) }");
        NotebookCell foreign = notebook.AddCodeCell("throw 'must never run'", "python");
        foreign.ExecutionCount = 27;
        foreign.AddOutput(NotebookOutput.CreateStream("stdout", "preserved"));
        NotebookCell result = notebook.AddCodeCell("Next-Number", "pwsh");
        result.AddOutput(NotebookOutput.CreateStream("stdout", "old"));
        NotebookCell blank = notebook.AddCodeCell(string.Empty);
        using PowerShellNotebookSession session = await StartAsync();

        NotebookExecutionReport first = await session.ExecuteAsync(notebook, new NotebookExecutionOptions { CellIndices = new[] { 2, 1, 0, 1 } });
        Assert.True(first.Succeeded);
        Assert.Equal(new[] { 0, 2 }, first.SkippedCellIndices);
        Assert.Equal(1, Assert.Single(first.Cells).ExecutionCount);

        NotebookExecutionReport second = await session.ExecuteAsync(notebook, new NotebookExecutionOptions { CellIndices = new[] { 4, 3 } });
        Assert.True(second.Succeeded);
        Assert.Equal(new[] { 2, 3 }, second.Cells.Select(cell => cell.ExecutionCount));
        Assert.Empty(blank.Outputs);
        Assert.Equal(1, setup.ExecutionCount);
        Assert.Equal(2, result.ExecutionCount);
        using NotebookDocument saved = NotebookDocument.Load(notebook.ToBytes());
        Assert.Equal("42", Assert.Single(saved.Cells[3].Outputs).Data["text/plain"].GetString());
        Assert.Equal(27, saved.Cells[2].ExecutionCount);
        Assert.Equal("preserved", Assert.Single(saved.Cells[2].Outputs).Text);
    }

    [PowerShellRuntimeFact]
    public async Task StreamsAndErrorsAreSavedInOrderAndCanContinueWithinTheSession() {
        using var notebook = NotebookDocument.Create();
        NotebookCell error = notebook.AddCodeCell("$n = 5; Write-Host 'begin'; Write-Warning 'warning'; $n; throw 'failed'", "ps1");
        NotebookCell after = notebook.AddCodeCell("$n + 1", "ps1");
        using PowerShellNotebookSession session = await StartAsync();
        NotebookExecutionReport report = await session.ExecuteAsync(notebook, new NotebookExecutionOptions { ContinueOnError = true });

        Assert.False(report.Succeeded);
        Assert.Equal(new[] { NotebookCellExecutionStatus.Failed, NotebookCellExecutionStatus.Succeeded },
            report.Cells.Select(cell => cell.Status));
        Assert.Collection(error.Outputs,
            item => { Assert.Equal("stdout", item.Name); Assert.Contains("begin", item.Text, StringComparison.Ordinal); },
            item => { Assert.Equal("stderr", item.Name); Assert.Contains("warning", item.Text, StringComparison.Ordinal); },
            item => Assert.Equal("5", item.Data["text/plain"].GetString()),
            item => { Assert.Equal(NotebookOutputKind.Error, item.Kind); Assert.Contains("failed", item.ErrorValue!, StringComparison.Ordinal); });
        Assert.Equal("6", Assert.Single(after.Outputs).Data["text/plain"].GetString());
        Assert.True(session.IsUsable);
    }

    [PowerShellRuntimeFact]
    public async Task TimeoutKeepsSavedPartialOutputAndStopsLaterCells() {
        using var notebook = NotebookDocument.Create();
        NotebookCell current = notebook.AddCodeCell("'before timeout'; Start-Sleep -Seconds 30", "powershell");
        NotebookCell later = notebook.AddCodeCell("'untouched'", "powershell");
        using PowerShellNotebookSession session = await StartAsync();

        NotebookExecutionReport report = await session.ExecuteAsync(notebook,
            new NotebookExecutionOptions { CellTimeout = TimeSpan.FromSeconds(3), ContinueOnError = true });

        Assert.Equal(NotebookCellExecutionStatus.TimedOut, Assert.Single(report.Cells).Status);
        Assert.Collection(current.Outputs,
            item => Assert.Equal("before timeout", item.Data["text/plain"].GetString()),
            item => Assert.Equal("NotebookCellTimeout", item.ErrorName));
        Assert.Empty(later.Outputs);
        Assert.Null(later.ExecutionCount);
        Assert.False(session.IsUsable);
    }

    [PowerShellRuntimeFact]
    public async Task CancellationFromAnOutputCallbackInterruptsTheLiveCell() {
        using var notebook = NotebookDocument.Create();
        NotebookCell cell = notebook.AddCodeCell("'before cancellation'; Start-Sleep -Seconds 30", "powershell");
        using PowerShellNotebookSession session = await StartAsync();
        using var cancellation = new CancellationTokenSource(TimeSpan.FromSeconds(15));
        NotebookExecutionEvent? observed = null;
        var options = new NotebookExecutionOptions {
            OutputReceived = output => {
                observed = output;
                cancellation.Cancel();
            }
        };

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => session.ExecuteAsync(notebook, options, cancellation.Token));

        Assert.NotNull(observed);
        Assert.Equal(cell.Id, observed.CellId);
        Assert.Collection(cell.Outputs,
            item => Assert.Equal("before cancellation", item.Data["text/plain"].GetString()),
            item => Assert.Equal("NotebookExecutionCanceled", item.ErrorName));
        Assert.False(session.IsUsable);
    }

    [PowerShellRuntimeFact]
    public async Task OutputLimitsDoNotDiscardRunspaceState() {
        using var notebook = NotebookDocument.Create();
        NotebookCell bounded = notebook.AddCodeCell("'abcd'; 'excess'; $later = 19", "powershell");
        NotebookCell next = notebook.AddCodeCell("$later", "powershell");
        using PowerShellNotebookSession session = await StartAsync();

        NotebookExecutionReport report = await session.ExecuteAsync(notebook,
            new NotebookExecutionOptions { OutputCharactersLimit = 3, OutputRecordsLimit = 1 });

        Assert.True(report.Succeeded);
        Assert.True(report.Cells[0].OutputTruncated);
        Assert.Equal("abc", Assert.Single(bounded.Outputs).Data["text/plain"].GetString());
        Assert.Equal("19", Assert.Single(next.Outputs).Data["text/plain"].GetString());
    }

    private static Task<PowerShellNotebookSession> StartAsync() =>
        PowerShellNotebookSession.StartAsync(new PowerShellNotebookSessionOptions { ExecutablePath = PowerShellRuntimeFactAttribute.Executable });

    private sealed class PowerShellRuntimeFactAttribute : FactAttribute {
        internal static string? Executable { get; } = FindPowerShell();

        public PowerShellRuntimeFactAttribute() {
            if (Executable == null) Skip = "PowerShell is required to exercise the notebook execution process.";
        }

        private static string? FindPowerShell() {
            string? configured = Environment.GetEnvironmentVariable("OFFICEIMO_TEST_POWERSHELL");
            if (!string.IsNullOrEmpty(configured)) return configured;
            bool isWindows = Path.DirectorySeparatorChar == '\\';
            string file = isWindows ? "pwsh.exe" : "pwsh";
            string? path = (Environment.GetEnvironmentVariable("PATH") ?? string.Empty).Split(Path.PathSeparator)
                .Where(static part => !string.IsNullOrWhiteSpace(part))
                .Select(part => Path.Combine(part.Trim('"'), file)).FirstOrDefault(File.Exists);
            if (path != null || !isWindows) return path;
            string legacy = Path.Combine(Environment.SystemDirectory, "WindowsPowerShell", "v1.0", "powershell.exe");
            return File.Exists(legacy) ? legacy : null;
        }
    }
}
