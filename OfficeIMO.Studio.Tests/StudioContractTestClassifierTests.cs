using System.Diagnostics;
using System.Security;
using Xunit;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioContractTestClassifierTests {
    [Fact]
    public void SuppressionRequiresACompleteRunWithOnlyTheKnownTeardownFailure() {
        if (!OperatingSystem.IsWindows()) return;

        string script = FindRepositoryFile("Build", "Invoke-StudioContractTests.ps1");
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO-Studio-TRX-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            var complete = RunClassifier(script, WriteTrx(directory, "complete.trx", total: 2, executed: 2,
                passed: 1, failed: 1, completed: 2, runLevelError: false));
            Assert.True(complete.ExitCode == 0, complete.Output);
            var incomplete = RunClassifier(script, WriteTrx(directory, "incomplete.trx", total: 2, executed: 1,
                passed: 0, failed: 1, completed: 1, runLevelError: false));
            Assert.True(incomplete.ExitCode != 0, incomplete.Output);
            var runError = RunClassifier(script, WriteTrx(directory, "run-error.trx", total: 2, executed: 2,
                passed: 1, failed: 1, completed: 2, runLevelError: true));
            Assert.True(runError.ExitCode != 0, runError.Output);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    private static string WriteTrx(string directory, string name, long total, long executed,
        long passed, long failed, long completed, bool runLevelError) {
        string path = Path.Combine(directory, name);
        string passedResult = passed == 0 ? string.Empty :
            "<UnitTestResult testName=\"Passing contract\" outcome=\"Passed\" />";
        string runProblem = runLevelError
            ? "<RunInfo outcome=\"Error\"><Text>Test host terminated unexpectedly.</Text></RunInfo>"
            : string.Empty;
        string knownMessage = SecurityElement.Escape(
            "System.NullReferenceException : Object reference not set to an instance of an object.")!;
        string knownStack = SecurityElement.Escape(
            "at Avalonia.Headless.HeadlessUnitTestSession.Dispose()\n   at Xunit.Runner.Run()")!;
        string xml = $"""
            <?xml version="1.0" encoding="utf-8"?>
            <TestRun xmlns="http://microsoft.com/schemas/VisualStudio/TeamTest/2010">
              <Times creation="2026-09-13T00:00:00Z" start="2026-09-13T00:00:00Z" finish="2026-09-13T00:00:01Z" />
              <Results>
                {passedResult}
                <UnitTestResult testName="Known teardown" outcome="Failed">
                  <Output><ErrorInfo>
                    <Message>{knownMessage}</Message>
                    <StackTrace>{knownStack}</StackTrace>
                  </ErrorInfo></Output>
                </UnitTestResult>
              </Results>
              <ResultSummary outcome="Failed">
                <Counters total="{total}" executed="{executed}" passed="{passed}" failed="{failed}"
                  error="0" timeout="0" aborted="0" inconclusive="0" passedButRunAborted="0"
                  notRunnable="0" notExecuted="0" disconnected="0" warning="0" completed="{completed}"
                  inProgress="0" pending="0" />
                <RunInfos>
                  <RunInfo outcome="Error"><Text>[xUnit.net] Known teardown [FAIL]</Text></RunInfo>
                  {runProblem}
                </RunInfos>
              </ResultSummary>
            </TestRun>
            """;
        File.WriteAllText(path, xml);
        return path;
    }

    private static (int ExitCode, string Output) RunClassifier(string script, string trxPath) {
        var startInfo = new ProcessStartInfo("pwsh") {
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };
        startInfo.ArgumentList.Add("-NoProfile");
        startInfo.ArgumentList.Add("-File");
        startInfo.ArgumentList.Add(script);
        startInfo.ArgumentList.Add("-TrxPath");
        startInfo.ArgumentList.Add(trxPath);
        startInfo.ArgumentList.Add("-Platform");
        startInfo.ArgumentList.Add("Windows");
        using var process = Process.Start(startInfo) ?? throw new InvalidOperationException("Could not start PowerShell.");
        if (!process.WaitForExit(30_000)) {
            process.Kill(entireProcessTree: true);
            throw new TimeoutException("Studio TRX classifier did not finish within 30 seconds.");
        }
        return (process.ExitCode, process.StandardOutput.ReadToEnd() + process.StandardError.ReadToEnd());
    }

    private static string FindRepositoryFile(params string[] parts) {
        DirectoryInfo? directory = new(AppContext.BaseDirectory);
        while (directory is not null) {
            string candidate = Path.Combine(new[] { directory.FullName }.Concat(parts).ToArray());
            if (File.Exists(candidate)) return candidate;
            directory = directory.Parent;
        }
        throw new FileNotFoundException("Could not locate the OfficeIMO repository file.", Path.Combine(parts));
    }
}
