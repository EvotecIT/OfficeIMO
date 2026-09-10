using System.Text;
using System.Diagnostics;
using System.Globalization;

namespace OfficeIMO.Invoicing.Validation.Tests;

public sealed class UnixInvoiceStandardsTheoryAttribute : TheoryAttribute {
    public UnixInvoiceStandardsTheoryAttribute() {
        if (OperatingSystem.IsWindows()) Skip = "Unix workspace permissions require a Unix filesystem.";
        else if (Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_STANDARDS_TESTS") != "1")
            Skip = "Uses the pinned Saxon artifact; run Build/Test-InvoicingStandards.ps1 on Unix.";
    }
}

public class InvoiceWorkspaceTests {
    [UnixInvoiceStandardsTheory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task FailedRuleProcessReportsIdentityButCompilerOnlyExecutionDoesNot(bool compile) {
        string probe = Directory.CreateTempSubdirectory("OfficeIMO.InvoiceProbe-").FullName;
        string executable = Path.Combine(probe, "runner");
        try {
            await File.WriteAllTextAsync(executable, "#!/bin/sh\nexit 7\n", new UTF8Encoding(false));
            if (!OperatingSystem.IsWindows()) File.SetUnixFileMode(executable, UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);
            var runner = new SaxonInvoiceRulesRunner(Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_SAXON_JAR")!, executable);
            bool started = false;
            await Assert.ThrowsAsync<InvalidOperationException>(() => runner.RunAsync(Encoding.UTF8.GetBytes("invoice"), Encoding.UTF8.GetBytes("rules"), compile,
                new Dictionary<string, InvoiceDiagnosticSeverity>(), CancellationToken.None, () => started = true));
            Assert.Equal(!compile, started);
        } finally { Directory.Delete(probe, recursive: true); }
    }

    [UnixInvoiceStandardsTheory]
    [InlineData("success")]
    [InlineData("failure")]
    [InlineData("cancellation")]
    [InlineData("timeout")]
    [InlineData("orphan-cancellation")]
    [InlineData("orphan-timeout")]
    public async Task WorkspaceIsPrivateWhileRunningAndRemovedAfterEveryExit(string outcome) {
        bool orphan = outcome.StartsWith("orphan-", StringComparison.Ordinal);
        string probe = Directory.CreateTempSubdirectory("OfficeIMO.InvoiceProbe-").FullName;
        string executable = Path.Combine(probe, "runner");
        string script = string.Join("\n", new[] {
            "#!/bin/sh", "set -eu",
            "printf '%s' \"$$\" > \"$0.parent\"",
            "printf '%s' \"$PWD\" > \"$0.workspace.tmp\"",
            "mv \"$0.workspace.tmp\" \"$0.workspace\"",
            "while [ ! -f \"$0.release\" ]; do sleep 0.02; done",
            outcome == "failure" ? "exit 7" : ":",
            orphan ? "sleep 10 & printf '%s' \"$!\" > \"$0.child\"" : ":",
            "for argument do case \"$argument\" in -o:*) output=${argument#-o:};; esac; done",
            "printf '%s' '<s:schematron-output xmlns:s=\"http://purl.oclc.org/dsdl/svrl\"><s:fired-rule context=\"Invoice\"/></s:schematron-output>' > \"$output\""
        });
        await File.WriteAllTextAsync(executable, script, new UTF8Encoding(false));
        if (!OperatingSystem.IsWindows()) File.SetUnixFileMode(executable, UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);
        using var cancellation = new CancellationTokenSource();
        Task<IReadOnlyList<InvoiceDiagnostic>>? running = null;
        Process? child = null;
        try {
            var runner = new SaxonInvoiceRulesRunner(Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_SAXON_JAR")!, executable, TimeSpan.FromSeconds(orphan ? 2 : 5));
            running = runner.RunAsync(Encoding.UTF8.GetBytes("private invoice"), Encoding.UTF8.GetBytes("rules"), false,
                new Dictionary<string, InvoiceDiagnosticSeverity>(), cancellation.Token);
            using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(10));
            while (!File.Exists(executable + ".workspace")) {
                if (running.IsCompleted) await running;
                await Task.Delay(20, deadline.Token);
            }
            string workspace = await File.ReadAllTextAsync(executable + ".workspace");
            if (!OperatingSystem.IsWindows())
                Assert.Equal(UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute, File.GetUnixFileMode(workspace));
            Assert.Equal("private invoice", await File.ReadAllTextAsync(Path.Combine(workspace, "invoice.xml")));
            if (orphan) {
                using Process parent = Process.GetProcessById(int.Parse(await File.ReadAllTextAsync(executable + ".parent"), CultureInfo.InvariantCulture));
                await File.WriteAllTextAsync(executable + ".release", "continue");
                await parent.WaitForExitAsync(deadline.Token);
                child = Process.GetProcessById(int.Parse(await File.ReadAllTextAsync(executable + ".child"), CultureInfo.InvariantCulture));
                Assert.False(child.HasExited);
            }
            if (outcome.EndsWith("cancellation", StringComparison.Ordinal)) {
                cancellation.Cancel();
                await Assert.ThrowsAnyAsync<OperationCanceledException>(() => running.WaitAsync(TimeSpan.FromSeconds(5)));
            } else if (outcome.EndsWith("timeout", StringComparison.Ordinal)) {
                TimeoutException error = await Assert.ThrowsAsync<TimeoutException>(() => running.WaitAsync(TimeSpan.FromSeconds(10)));
                Assert.Contains("Saxon exceeded", error.Message);
            } else {
                await File.WriteAllTextAsync(executable + ".release", "continue");
                if (outcome == "failure") await Assert.ThrowsAsync<InvalidOperationException>(() => running);
                else Assert.Empty(await running);
            }
            Assert.False(Directory.Exists(workspace));
        } finally {
            cancellation.Cancel();
            if (child != null) {
                try { if (!child.HasExited) child.Kill(); await child.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(5)); }
                finally { child.Dispose(); }
            }
            if (running != null) { try { await running; } catch { } }
            Directory.Delete(probe, recursive: true);
        }
    }
}
