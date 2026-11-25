using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
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
    public async Task EscapedTextSurvivesExecutionAndNotebookSave() {
        const string text = "\"quoted\"\\path\t\b\fline\nreturn\rnull\0\u001b\u2028\u2029\U0001F680";
        string encoded = Convert.ToBase64String(Encoding.UTF8.GetBytes(text));
        using var notebook = NotebookDocument.Create();
        notebook.AddCodeCell(
            $"$output = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('{encoded}')); $output; Write-Host -NoNewline -Object $output",
            "powershell");
        using PowerShellNotebookSession session = await StartAsync();

        NotebookExecutionReport report = await session.ExecuteAsync(notebook);

        Assert.Equal(NotebookCellExecutionStatus.Succeeded, Assert.Single(report.Cells).Status);
        using NotebookDocument saved = NotebookDocument.Load(notebook.ToBytes());
        foreach (NotebookDocument document in new[] { notebook, saved }) {
            Assert.Collection(Assert.Single(document.Cells).Outputs,
                output => {
                    Assert.Equal(NotebookOutputKind.ExecuteResult, output.Kind);
                    Assert.Equal(text, output.Data["text/plain"].GetString());
                },
                output => {
                    Assert.Equal(NotebookOutputKind.Stream, output.Kind);
                    Assert.Equal("stdout", output.Name);
                    Assert.Equal(text, output.Text);
                });
        }
    }

    [PowerShellRuntimeFact]
    public async Task TimeoutKeepsSavedPartialOutputAndStopsLaterCells() {
        using var notebook = NotebookDocument.Create();
        NotebookCell current = notebook.AddCodeCell("'before timeout'; 'omitted'; Start-Sleep -Seconds 30", "powershell");
        NotebookCell later = notebook.AddCodeCell("'untouched'", "powershell");
        using PowerShellNotebookSession session = await StartAsync();

        NotebookExecutionReport report = await session.ExecuteAsync(notebook,
            new NotebookExecutionOptions { CellTimeout = TimeSpan.FromSeconds(3), ContinueOnError = true, OutputRecordsLimit = 1 });

        NotebookCellExecutionResult result = Assert.Single(report.Cells);
        Assert.Equal(NotebookCellExecutionStatus.TimedOut, result.Status);
        Assert.True(result.OutputTruncated);
        Assert.Collection(current.Outputs,
            item => Assert.Equal("before timeout", item.Data["text/plain"].GetString()),
            item => Assert.Equal("NotebookCellTimeout", item.ErrorName));
        Assert.Empty(later.Outputs);
        Assert.Null(later.ExecutionCount);
        Assert.False(session.IsUsable);
    }

    [PowerShellRuntimeFact]
    public Task CancellationFromAnOutputCallbackStopsTheCellAndItsNativeChild() =>
        AssertCancellationStopsNativeChildAsync("""
            $executable = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
            $child = Start-Process -FilePath $executable -NoNewWindow -ArgumentList '-NoProfile','-NonInteractive','-Command','Start-Sleep -Seconds 30' -PassThru
            $child.Id
            Start-Sleep -Seconds 30
            """);

    [UnixProcessGroupFact]
    public Task CancellationStopsBackgroundChildAfterItsShellExits() =>
        AssertCancellationStopsNativeChildAsync("""
            & /bin/sh -c '/bin/sleep 30 </dev/null >/dev/null 2>&1 & echo "$$,$!"'
            Start-Sleep -Seconds 30
            """, expectExitedShell: true);

    private static async Task AssertCancellationStopsNativeChildAsync(string source, bool expectExitedShell = false) {
        using var notebook = NotebookDocument.Create();
        NotebookCell cell = notebook.AddCodeCell(source, "powershell");
        using PowerShellNotebookSession session = await StartAsync();
        using var cancellation = new CancellationTokenSource(TimeSpan.FromSeconds(15));
        NotebookExecutionEvent? observed = null;
        Process? child = null;
        var options = new NotebookExecutionOptions {
            OutputReceived = output => {
                observed = output;
                string[] processIds = output.Output.Data["text/plain"].GetString()!.Split(',');
                Assert.Equal(expectExitedShell ? 2 : 1, processIds.Length);
                int pid = int.Parse(processIds[processIds.Length - 1], CultureInfo.InvariantCulture);
                child = Process.GetProcessById(pid);
                if (expectExitedShell) {
                    int shellId = int.Parse(processIds[0], CultureInfo.InvariantCulture);
                    Assert.True(SpinWait.SpinUntil(() => HasShellExited(shellId), TimeSpan.FromSeconds(5)),
                        "The background child's shell must exit before cancellation starts.");
                }
                Assert.False(HasStopped(child), "The notebook child must be running when cancellation starts.");
                cancellation.Cancel();
            }
        };

        try {
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => session.ExecuteAsync(notebook, options, cancellation.Token));

            Assert.NotNull(observed);
            Assert.NotNull(child);
            Assert.Equal(cell.Id, observed.CellId);
            Assert.Collection(cell.Outputs,
                item => Assert.Equal(observed.Output.Data["text/plain"].GetString(), item.Data["text/plain"].GetString()),
                item => Assert.Equal("NotebookExecutionCanceled", item.ErrorName));
            Assert.True(SpinWait.SpinUntil(() => HasStopped(child), TimeSpan.FromSeconds(5)), "The executed cell's native child is still running.");
            Assert.False(session.IsUsable);
        } finally {
            using (child) {
                if (child != null && !HasStopped(child)) {
                    try {
                        child.Kill();
                    } catch (InvalidOperationException) when (HasStopped(child)) {
                        // The child exited between the status check and the cleanup signal.
                    }
                    Assert.True(SpinWait.SpinUntil(() => HasStopped(child), TimeSpan.FromSeconds(5)), "The test child survived cleanup.");
                }
            }
        }
    }

    private static bool HasShellExited(int pid) {
        try {
            using Process shell = Process.GetProcessById(pid);
            return HasStopped(shell);
        } catch (ArgumentException) {
            return true;
        }
    }

    private static bool HasStopped(Process process) {
        if (process.HasExited) return true;
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Linux)) return false;
        try {
            // A test container's PID 1 may leave a terminated orphan in zombie state.
            string status = File.ReadAllText("/proc/" + process.Id.ToString(CultureInfo.InvariantCulture) + "/stat");
            int state = status.LastIndexOf(')') + 2;
            return state < status.Length ? status[state] == 'Z' : process.HasExited;
        } catch (IOException) {
            return process.HasExited;
        }
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

    [PowerShellRuntimeFact]
    public async Task OmittedCommandErrorStillStopsTheNotebook() {
        using var notebook = NotebookDocument.Create();
        NotebookCell failed = notebook.AddCodeCell("'keep'; 'omit'; Get-OfficeImoNoSuchNotebookCommand", "powershell");
        NotebookCell later = notebook.AddCodeCell("'untouched'", "powershell");
        using PowerShellNotebookSession session = await StartAsync();

        NotebookExecutionReport report = await session.ExecuteAsync(notebook,
            new NotebookExecutionOptions { OutputRecordsLimit = 1 });

        NotebookCellExecutionResult result = Assert.Single(report.Cells);
        Assert.Equal(NotebookCellExecutionStatus.Failed, result.Status);
        Assert.True(result.OutputTruncated);
        Assert.Equal("keep", Assert.Single(failed.Outputs).Data["text/plain"].GetString());
        Assert.Null(later.ExecutionCount);
        Assert.True(session.IsUsable);
    }

    private static Task<PowerShellNotebookSession> StartAsync() =>
        PowerShellNotebookSession.StartAsync(new PowerShellNotebookSessionOptions { ExecutablePath = PowerShellRuntimeFactAttribute.Executable });

    private sealed class UnixProcessGroupFactAttribute : PowerShellRuntimeFactAttribute {
        public UnixProcessGroupFactAttribute() {
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows) ||
                !(File.Exists("/usr/bin/setsid") || File.Exists("/bin/setsid") ||
                  (RuntimeInformation.IsOSPlatform(OSPlatform.OSX) && File.Exists("/usr/bin/perl")))) {
                Skip ??= "A Unix session launcher is required to exercise reparented process cleanup.";
            }
        }
    }

    private class PowerShellRuntimeFactAttribute : FactAttribute {
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
