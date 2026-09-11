using System.Text;

namespace OfficeIMO.Invoicing.Validation.Tests;

public partial class InvoiceStandardsTests {
    [InvoiceStandardsTheory]
    [InlineData("saxon-he-12.10.jar", false)]
    [InlineData("saxon-he-12.10.jar", true)]
    [InlineData("lib/xmlresolver-5.3.3.jar", false)]
    [InlineData("lib/xmlresolver-5.3.3.jar", true)]
    [InlineData("lib/xmlresolver-5.3.3-data.jar", false)]
    [InlineData("lib/xmlresolver-5.3.3-data.jar", true)]
    [InlineData("lib/jline-2.14.6.jar", false)]
    [InlineData("lib/jline-2.14.6.jar", true)]
    public async Task RuntimeFilesAreVerifiedAtConstructionAndBeforeExecution(string relativePath, bool missing) {
        using var runtime = new RuntimeCopy();
        var runner = new SaxonInvoiceRulesRunner(runtime.Jar);
        string changed = Path.Combine(runtime.Directory, relativePath);
        if (missing) File.Delete(changed);
        else File.WriteAllText(changed, "corrupt runtime");
        if (missing) Assert.Throws<FileNotFoundException>(() => new SaxonInvoiceRulesRunner(runtime.Jar));
        else Assert.Throws<InvalidDataException>(() => new SaxonInvoiceRulesRunner(runtime.Jar));
        var validator = new InvoiceValidator(Bundle(), runner);
        InvoiceValidationReport report = await validator.ValidateAsync(InvoiceSerializer.Write(OfficeIMO.Invoicing.Tests.InvoiceFixture.Create()), InvoiceRulesRelease.En16931_1_3_16);
        Assert.Equal(InvoiceValidationStatus.Failed, report.BusinessRulesStatus);
        Assert.Null(report.Runner);
        Assert.Contains(report.Diagnostics, d => d.Code == "INV-RULES-ENGINE");
    }

    [InvoiceStandardsTheory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task RunningRulesUseAnIsolatedVerifiedRuntimeSnapshot(bool compile) {
        using var runtime = new RuntimeCopy();
        var runner = new SaxonInvoiceRulesRunner(runtime.Jar);
        byte[] source = Encoding.UTF8.GetBytes(compile
            ? "<schema xmlns='http://purl.oclc.org/dsdl/schematron' queryBinding='xslt2'><pattern><rule context='Invoice'><assert id='test' test='true()'>ok</assert></rule></pattern></schema>"
            : "<xsl:stylesheet xmlns:xsl='http://www.w3.org/1999/XSL/Transform' version='2.0'><xsl:template match='/'><schematron-output xmlns='http://purl.oclc.org/dsdl/svrl'><fired-rule context='Invoice'/></schematron-output></xsl:template></xsl:stylesheet>");
        IReadOnlyList<InvoiceDiagnostic> diagnostics = await runner.RunAsync(Encoding.UTF8.GetBytes("<Invoice/>"), source, compile,
            new Dictionary<string, InvoiceDiagnosticSeverity>(), CancellationToken.None,
            () => { foreach (string path in System.IO.Directory.EnumerateFiles(runtime.Directory, "*.jar", SearchOption.AllDirectories)) File.WriteAllText(path, "changed after verification"); });
        Assert.Empty(diagnostics);
    }

    private sealed class RuntimeCopy : IDisposable {
        internal string Directory { get; } = System.IO.Directory.CreateTempSubdirectory("OfficeIMO.InvoiceRuntimeTest-").FullName;
        internal string Jar => Path.Combine(Directory, "saxon-he-12.10.jar");
        internal RuntimeCopy() {
            try {
                string original = Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_SAXON_JAR")!;
                File.Copy(original, Jar);
                System.IO.Directory.CreateDirectory(Path.Combine(Directory, "lib"));
                foreach (string name in new[] { "xmlresolver-5.3.3.jar", "xmlresolver-5.3.3-data.jar", "jline-2.14.6.jar" })
                    File.Copy(Path.Combine(Path.GetDirectoryName(original)!, "lib", name), Path.Combine(Directory, "lib", name));
            } catch { Dispose(); throw; }
        }
        public void Dispose() => System.IO.Directory.Delete(Directory, true);
    }
}
