using System.Text;

namespace OfficeIMO.Invoicing.Validation.Tests;

public class InvoiceOutputLimitTests {
    [UnixInvoiceStandardsTheory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task OutputFileCannotExceedItsLimitWhileTheProducerIsRunning(bool compiler) {
        string probe = Directory.CreateTempSubdirectory("OfficeIMO.InvoiceOutputProbe-").FullName;
        string executable = Path.Combine(probe, "runner");
        string script = """
            #!/bin/sh
            set -eu
            printf '%s' "$PWD" > "$0.workspace.tmp"
            mv "$0.workspace.tmp" "$0.workspace"
            output=/dev/stdout
            for argument do case "$argument" in -o:*) output=${argument#-o:};; esac; done
            while [ ! -f "$0.release" ]; do sleep 0.02; done
            exec python3 - "$output" <<'PY'
            import sys, time
            with open(sys.argv[1], 'wb', buffering=0) as output:
                for n in range(320):
                    output.write(b'x' * 65536)
                time.sleep(30)
            PY
            """;
        await File.WriteAllTextAsync(executable, script.Replace("\r\n", "\n"), new UTF8Encoding(false));
        if (!OperatingSystem.IsWindows()) File.SetUnixFileMode(executable, UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);
        using var cancellation = new CancellationTokenSource(TimeSpan.FromSeconds(15));
        Task<IReadOnlyList<InvoiceDiagnostic>>? running = null;
        try {
            var runner = new SaxonInvoiceRulesRunner(Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_SAXON_JAR")!, executable, TimeSpan.FromSeconds(10));
            running = runner.RunAsync(Encoding.UTF8.GetBytes("invoice"), Encoding.UTF8.GetBytes("rules"), compiler,
                new Dictionary<string, InvoiceDiagnosticSeverity>(), cancellation.Token);
            while (!File.Exists(executable + ".workspace")) {
                if (running.IsCompleted) await running;
                await Task.Delay(10, cancellation.Token);
            }
            string workspace = await File.ReadAllTextAsync(executable + ".workspace", cancellation.Token);
            string output = Path.Combine(workspace, compiler ? "rules.xsl" : "report.xml");
            await File.WriteAllTextAsync(executable + ".release", "continue", cancellation.Token);
            while (!running.IsCompleted) {
                try { if (File.Exists(output)) Assert.InRange(new FileInfo(output).Length, 0, 16 * 1024 * 1024); }
                catch (FileNotFoundException) { /* The runner can remove its workspace between the existence and length checks. */ }
                catch (DirectoryNotFoundException) { }
                await Task.Delay(5, cancellation.Token);
            }
            InvalidDataException error = await Assert.ThrowsAsync<InvalidDataException>(() => running);
            Assert.Contains("16 MiB", error.Message);
            Assert.False(Directory.Exists(workspace));
        } finally {
            cancellation.Cancel();
            if (running != null) { try { await running; } catch { } }
            Directory.Delete(probe, recursive: true);
        }
    }

    [InvoiceStandardsTheory]
    [InlineData(200000)]
    public async Task RealSaxonStopsWhenItsReportExceedsTheOutputLimit(int assertions) {
        string stylesheet = """
            <xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:s="http://purl.oclc.org/dsdl/svrl" version="2.0">
              <xsl:template match="/">
                <s:schematron-output><s:fired-rule context="Invoice"/>
                  <xsl:for-each select="1 to COUNT">
                    <s:failed-assert id="TEST" location="Invoice"><s:text>Repeated diagnostic text for the bounded report output contract.</s:text></s:failed-assert>
                  </xsl:for-each>
                </s:schematron-output>
              </xsl:template>
            </xsl:stylesheet>
            """;
        var runner = new SaxonInvoiceRulesRunner(Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_SAXON_JAR")!, Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_JAVA") ?? "java");
        InvalidDataException error = await Assert.ThrowsAsync<InvalidDataException>(() => runner.RunAsync(Encoding.UTF8.GetBytes("<Invoice/>"),
            Encoding.UTF8.GetBytes(stylesheet.Replace("COUNT", assertions.ToString(System.Globalization.CultureInfo.InvariantCulture))), false,
            new Dictionary<string, InvoiceDiagnosticSeverity>(), CancellationToken.None));
        Assert.Contains("16 MiB", error.Message);
    }
}
