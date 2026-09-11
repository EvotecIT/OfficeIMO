using System.Diagnostics;
using System.Reflection;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Invoicing.Validation;

/// <summary>Runs pinned Schematron artifacts with a locally installed, explicitly configured SaxonJ-HE 12.10 runtime.</summary>
public sealed class SaxonInvoiceRulesRunner {
    /// <summary>Official SaxonJ-HE distribution used by this adapter; retain its lib directory next to the main JAR.</summary>
    public const string DownloadUrl = "https://downloads.saxonica.com/SaxonJ/HE/12/SaxonHE12-10J.zip";
    /// <summary>SHA-256 of the official SaxonJ-HE 12.10 distribution archive.</summary>
    public const string ArchiveSha256 = "1C7DB9F726DF835349C64EDD631DE0310ECA31291100230064EBA153F607B0BE";
    /// <summary>SHA-256 of saxon-he-12.10.jar.</summary>
    public const string JarSha256 = "89BCD071666C3268EE8C5E91E9C9812AD52B3E49D38A3A971720FDB428BFEEA9";
    private readonly string _jar;
    private readonly string _java;
    private readonly TimeSpan _timeout;
    /// <summary>Creates a runner without downloading or installing anything. Each subprocess has a bounded execution time.</summary>
    public SaxonInvoiceRulesRunner(string saxonJarPath, string javaExecutable = "java", TimeSpan? timeout = null) {
        ArgumentException.ThrowIfNullOrWhiteSpace(saxonJarPath); ArgumentException.ThrowIfNullOrWhiteSpace(javaExecutable);
        _jar = Path.GetFullPath(saxonJarPath); _java = javaExecutable; _timeout = timeout ?? TimeSpan.FromSeconds(60);
        if (_timeout < TimeSpan.FromSeconds(1) || _timeout > TimeSpan.FromMinutes(5)) throw new ArgumentOutOfRangeException(nameof(timeout));
        SaxonInvoiceRuntime.Load(_jar);
    }
    /// <summary>Verified main engine and companion JAR identities included in validation reports.</summary>
    public string Identity => SaxonInvoiceRuntime.Identity;

    internal async Task<IReadOnlyList<InvoiceDiagnostic>> RunAsync(byte[] xml, byte[] source, bool compile,
        IReadOnlyDictionary<string, InvoiceDiagnosticSeverity> overrides, CancellationToken cancellationToken, Action? ruleProcessStarted = null) {
        cancellationToken.ThrowIfCancellationRequested();
        SaxonInvoiceRuntime runtime = SaxonInvoiceRuntime.Load(_jar);
        string directory = Directory.CreateTempSubdirectory("OfficeIMO.InvoiceRules-").FullName;
        try {
            string runtimeJar = await runtime.WriteAsync(directory, cancellationToken).ConfigureAwait(false);
            string input = Path.Combine(directory, "invoice.xml"), stylesheet = Path.Combine(directory, "rules.xsl"), report = Path.Combine(directory, "report.xml");
            await File.WriteAllBytesAsync(input, xml, cancellationToken).ConfigureAwait(false);
            if (compile) {
                string schema = Path.Combine(directory, "rules.sch");
                await File.WriteAllBytesAsync(schema, source, cancellationToken).ConfigureAwait(false);
                foreach (string name in new[] { "iso_svrl_for_xslt2.xsl", "iso_schematron_skeleton_for_saxon.xsl" }) {
                    using Stream resource = Assembly.GetExecutingAssembly().GetManifestResourceStream("OfficeIMO.Invoicing.Validation.Compiler." + name)
                        ?? throw new InvalidDataException("Schematron compiler resource is missing.");
                    using var output = File.Create(Path.Combine(directory, name));
                    await resource.CopyToAsync(output, cancellationToken).ConfigureAwait(false);
                }
                await TransformAsync(schema, Path.Combine(directory, "iso_svrl_for_xslt2.xsl"), stylesheet, directory, runtimeJar, true, cancellationToken).ConfigureAwait(false);
            } else await File.WriteAllBytesAsync(stylesheet, source, cancellationToken).ConfigureAwait(false);
            await TransformAsync(input, stylesheet, report, directory, runtimeJar, false, cancellationToken, ruleProcessStarted).ConfigureAwait(false);
            using Stream result = File.OpenRead(report);
            return ReadSvrl(result, overrides);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    private async Task TransformAsync(string input, string stylesheet, string output, string workingDirectory, string runtimeJar, bool compiler, CancellationToken cancellationToken, Action? ruleProcessStarted = null) {
        var start = new ProcessStartInfo(_java) { UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true, WorkingDirectory = workingDirectory };
        foreach (string argument in new[] { "-Xmx256m", "-Djavax.xml.accessExternalDTD=", "-Djavax.xml.accessExternalSchema=", "-Djavax.xml.accessExternalStylesheet=file",
            "-jar", runtimeJar, "-s:" + input, "-xsl:" + stylesheet, "-dtd:off", "-xi:off", "-ext:off" }) start.ArgumentList.Add(argument);
        if (compiler) start.ArgumentList.Add("allow-foreign=true");
        // Omitting -o sends Saxon's principal result to stdout. Own every disk write so
        // neither compiled stylesheets nor reports can grow beyond the byte limit.
        await using var result = new FileStream(output, FileMode.CreateNew, FileAccess.Write, FileShare.Read, 8192, useAsync: true);
        using var process = new Process { StartInfo = start };
        if (!process.Start()) throw new InvalidOperationException("Saxon process could not start.");
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken); timeout.CancelAfter(_timeout);
        Task standardOutput = CopyResultAsync(process.StandardOutput.BaseStream, result, timeout.Token);
        Task<string> standardError = DrainAsync(process.StandardError, timeout.Token);
        Task completion = Task.WhenAll(CancelOnFailureAsync(standardOutput, timeout), CancelOnFailureAsync(standardError, timeout),
            CancelOnFailureAsync(process.WaitForExitAsync(timeout.Token), timeout));
        try {
            ruleProcessStarted?.Invoke();
            await completion.ConfigureAwait(false);
        } catch (Exception error) {
            timeout.Cancel();
            if (!process.HasExited) {
                try { process.Kill(entireProcessTree: true); }
                catch (InvalidOperationException) when (process.HasExited) { }
            }
            await process.WaitForExitAsync(CancellationToken.None).WaitAsync(TimeSpan.FromSeconds(10)).ConfigureAwait(false);
            try { await completion.WaitAsync(TimeSpan.FromSeconds(10)).ConfigureAwait(false); }
            catch (Exception) { /* Observe cancelled or failed pipe tasks before disposing their streams. */ }
            cancellationToken.ThrowIfCancellationRequested();
            if (error is OperationCanceledException) throw new TimeoutException("Saxon exceeded the configured " + _timeout.TotalSeconds + " second timeout.");
            throw;
        }
        if (process.ExitCode != 0) throw new InvalidOperationException("Saxon exited with code " + process.ExitCode + ": " + await standardError.ConfigureAwait(false));
        if (result.Length == 0) throw new InvalidDataException("Saxon produced no output.");
    }
    private static async Task CancelOnFailureAsync(Task operation, CancellationTokenSource cancellation) {
        try { await operation.ConfigureAwait(false); }
        catch { cancellation.Cancel(); throw; }
    }
    private static async Task CopyResultAsync(Stream input, Stream output, CancellationToken cancellationToken) {
        const int maximumBytes = 16 * 1024 * 1024;
        byte[] buffer = new byte[8192]; int total = 0, count;
        while ((count = await input.ReadAsync(buffer.AsMemory(), cancellationToken).ConfigureAwait(false)) != 0) {
            if (count > maximumBytes - total) throw new InvalidDataException("Schematron output exceeds 16 MiB.");
            await output.WriteAsync(buffer.AsMemory(0, count), cancellationToken).ConfigureAwait(false);
            total += count;
        }
    }
    private static async Task<string> DrainAsync(StreamReader reader, CancellationToken cancellationToken) {
        var result = new StringBuilder(); char[] buffer = new char[4096]; int count;
        while ((count = await reader.ReadAsync(buffer.AsMemory(), cancellationToken).ConfigureAwait(false)) != 0) {
            int remaining = 65536 - result.Length;
            if (remaining > 0) result.Append(buffer, 0, Math.Min(remaining, count));
        }
        return result.ToString();
    }
    internal static IReadOnlyList<InvoiceDiagnostic> ReadSvrl(Stream input, IReadOnlyDictionary<string, InvoiceDiagnosticSeverity> overrides) {
        using XmlReader reader = XmlReader.Create(input, InvoiceRuleBundle.XmlSettings());
        XDocument document = XDocument.Load(reader);
        XNamespace svrl = "http://purl.oclc.org/dsdl/svrl";
        if (document.Root?.Name != svrl + "schematron-output" || !document.Descendants(svrl + "fired-rule").Any())
            throw new InvalidDataException("The rule engine did not return an executed SVRL report.");
        var diagnostics = new InvoiceDiagnosticBuffer();
        foreach (XElement element in document.Descendants().Where(e => e.Name == svrl + "failed-assert" || e.Name == svrl + "successful-report")) {
            string code = (string?)element.Attribute("id") ?? "SCHEMATRON";
            string text = string.Join(" ", element.Elements(svrl + "text").Select(e => e.Value.Trim()));
            string? flag = (string?)element.Attribute("flag") ?? (string?)element.Attribute("role");
            InvoiceDiagnosticSeverity severity = overrides.TryGetValue(code, out InvoiceDiagnosticSeverity value) ? value : InvoiceRuleBundle.ParseSeverity(flag);
            diagnostics.Add(code, text, (string?)element.Attribute("location") ?? "Invoice", severity);
        }
        return diagnostics.ToList().AsReadOnly();
    }
}
