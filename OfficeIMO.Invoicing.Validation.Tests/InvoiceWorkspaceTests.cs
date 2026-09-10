using System.Text;

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
    [InlineData("success")]
    [InlineData("failure")]
    [InlineData("cancellation")]
    [InlineData("timeout")]
    public async Task WorkspaceIsPrivateWhileRunningAndRemovedAfterEveryExit(string outcome) {
        string probe = Directory.CreateTempSubdirectory("OfficeIMO.InvoiceProbe-").FullName;
        string executable = Path.Combine(probe, "runner");
        string script = string.Join("\n", new[] {
            "#!/bin/sh", "set -eu",
            "printf '%s' \"$PWD\" > \"$0.workspace.tmp\"",
            "mv \"$0.workspace.tmp\" \"$0.workspace\"",
            "while [ ! -f \"$0.release\" ]; do sleep 0.02; done",
            outcome == "failure" ? "exit 7" : ":",
            "for argument do case \"$argument\" in -o:*) output=${argument#-o:};; esac; done",
            "printf '%s' '<s:schematron-output xmlns:s=\"http://purl.oclc.org/dsdl/svrl\"><s:fired-rule context=\"Invoice\"/></s:schematron-output>' > \"$output\""
        });
        await File.WriteAllTextAsync(executable, script, new UTF8Encoding(false));
        if (!OperatingSystem.IsWindows()) File.SetUnixFileMode(executable, UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);
        using var cancellation = new CancellationTokenSource();
        Task<IReadOnlyList<InvoiceDiagnostic>>? running = null;
        try {
            var runner = new SaxonInvoiceRulesRunner(Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_SAXON_JAR")!, executable, TimeSpan.FromSeconds(5));
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
            if (outcome == "cancellation") {
                cancellation.Cancel();
                await Assert.ThrowsAnyAsync<OperationCanceledException>(() => running);
            } else if (outcome == "timeout") {
                await Assert.ThrowsAsync<TimeoutException>(() => running);
            } else {
                await File.WriteAllTextAsync(executable + ".release", "continue");
                if (outcome == "failure") await Assert.ThrowsAsync<InvalidOperationException>(() => running);
                else Assert.Empty(await running);
            }
            Assert.False(Directory.Exists(workspace));
        } finally {
            cancellation.Cancel();
            if (running != null) { try { await running; } catch { } }
            Directory.Delete(probe, recursive: true);
        }
    }
}
