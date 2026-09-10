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
        InvoiceRuleBundle.ReadPinned(_jar, JarSha256, 8 * 1024 * 1024);
    }
    /// <summary>Verified main engine identity included in validation reports.</summary>
    public string Identity => "SaxonJ-HE 12.10; SHA256=" + JarSha256;

    internal async Task<IReadOnlyList<InvoiceDiagnostic>> RunAsync(byte[] xml, byte[] source, bool compile,
        IReadOnlyDictionary<string, InvoiceDiagnosticSeverity> overrides, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        InvoiceRuleBundle.ReadPinned(_jar, JarSha256, 8 * 1024 * 1024);
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.InvoiceRules-" + Guid.NewGuid().ToString("N"));
        if (Directory.Exists(directory)) throw new IOException("Validation workspace already exists.");
        Directory.CreateDirectory(directory);
        try {
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
                await TransformAsync(schema, Path.Combine(directory, "iso_svrl_for_xslt2.xsl"), stylesheet, directory, true, cancellationToken).ConfigureAwait(false);
            } else await File.WriteAllBytesAsync(stylesheet, source, cancellationToken).ConfigureAwait(false);
            await TransformAsync(input, stylesheet, report, directory, false, cancellationToken).ConfigureAwait(false);
            if (new FileInfo(report).Length > 16 * 1024 * 1024) throw new InvalidDataException("Schematron report exceeds 16 MiB.");
            using Stream result = File.OpenRead(report);
            return ReadSvrl(result, overrides);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    private async Task TransformAsync(string input, string stylesheet, string output, string workingDirectory, bool compiler, CancellationToken cancellationToken) {
        var start = new ProcessStartInfo(_java) { UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true, WorkingDirectory = workingDirectory };
        foreach (string argument in new[] { "-Xmx256m", "-Djavax.xml.accessExternalDTD=", "-Djavax.xml.accessExternalSchema=", "-Djavax.xml.accessExternalStylesheet=file",
            "-jar", _jar, "-s:" + input, "-xsl:" + stylesheet, "-o:" + output, "-dtd:off", "-xi:off", "-ext:off" }) start.ArgumentList.Add(argument);
        if (compiler) start.ArgumentList.Add("allow-foreign=true");
        using var process = new Process { StartInfo = start };
        if (!process.Start()) throw new InvalidOperationException("Saxon process could not start.");
        Task<string> standardOutput = DrainAsync(process.StandardOutput), standardError = DrainAsync(process.StandardError);
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken); timeout.CancelAfter(_timeout);
        try {
            await process.WaitForExitAsync(timeout.Token).ConfigureAwait(false);
        } catch (OperationCanceledException) {
            if (!process.HasExited) process.Kill(entireProcessTree: true);
            await process.WaitForExitAsync(CancellationToken.None).WaitAsync(TimeSpan.FromSeconds(10)).ConfigureAwait(false);
            await Task.WhenAll(standardOutput, standardError).WaitAsync(TimeSpan.FromSeconds(10)).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            throw new TimeoutException("Saxon exceeded the configured " + _timeout.TotalSeconds + " second timeout.");
        }
        string[] messages = await Task.WhenAll(standardOutput, standardError).ConfigureAwait(false);
        if (process.ExitCode != 0) throw new InvalidOperationException("Saxon exited with code " + process.ExitCode + ": " + string.Join(" ", messages));
        if (!File.Exists(output)) throw new InvalidDataException("Saxon produced no output file.");
    }
    private static async Task<string> DrainAsync(StreamReader reader) {
        var result = new StringBuilder(); char[] buffer = new char[4096]; int count;
        while ((count = await reader.ReadAsync(buffer).ConfigureAwait(false)) != 0) {
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
        var diagnostics = new List<InvoiceDiagnostic>();
        foreach (XElement element in document.Descendants().Where(e => e.Name == svrl + "failed-assert" || e.Name == svrl + "successful-report")) {
            string code = (string?)element.Attribute("id") ?? "SCHEMATRON";
            string text = string.Join(" ", element.Elements(svrl + "text").Select(e => e.Value.Trim()));
            string? flag = (string?)element.Attribute("flag") ?? (string?)element.Attribute("role");
            InvoiceDiagnosticSeverity severity = overrides.TryGetValue(code, out InvoiceDiagnosticSeverity value) ? value : InvoiceRuleBundle.ParseSeverity(flag);
            diagnostics.Add(new InvoiceDiagnostic(code, text, (string?)element.Attribute("location") ?? "Invoice", severity));
            if (diagnostics.Count > 1000) throw new InvalidDataException("Schematron exceeded 1,000 diagnostics.");
        }
        return diagnostics.AsReadOnly();
    }
}
