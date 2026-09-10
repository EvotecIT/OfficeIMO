using System.Text;

namespace OfficeIMO.Invoicing.Validation.Tests;

public class InvoiceValidationBoundaryTests {
    [Theory]
    [InlineData("warning", InvoiceDiagnosticSeverity.Warning)]
    [InlineData("fatal", InvoiceDiagnosticSeverity.Error)]
    public void ExcessiveSvrlDiagnosticsRetainLateSeverity(string finalFlag, InvoiceDiagnosticSeverity expected) {
        var xml = new StringBuilder("<s:schematron-output xmlns:s='http://purl.oclc.org/dsdl/svrl'><s:fired-rule context='Invoice'/>");
        for (int index = 0; index < 1100; index++) xml.Append("<s:failed-assert id='warning' flag='warning'><s:text>Warning</s:text></s:failed-assert>");
        xml.Append("<s:failed-assert id='last' flag='").Append(finalFlag).Append("'><s:text>Final finding</s:text></s:failed-assert></s:schematron-output>");
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(xml.ToString()));
        IReadOnlyList<InvoiceDiagnostic> result = SaxonInvoiceRulesRunner.ReadSvrl(stream, new Dictionary<string, InvoiceDiagnosticSeverity>());
        Assert.Equal(1000, result.Count);
        InvoiceDiagnostic summary = Assert.Single(result.Where(d => d.Code == "INV-DIAGNOSTICS-TRUNCATED"));
        Assert.Equal(expected, summary.Severity);
    }

    [Fact]
    public void EmptySvrlDoesNotCountAsExecutedRules() {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("<svrl:schematron-output xmlns:svrl='http://purl.oclc.org/dsdl/svrl'/>"));
        Assert.Throws<InvalidDataException>(() => SaxonInvoiceRulesRunner.ReadSvrl(stream, new Dictionary<string, InvoiceDiagnosticSeverity>()));
    }

    [Fact]
    public void SeverityOverridesRemainSeparateFromRawRuleFlags() {
        const string svrl = "<s:schematron-output xmlns:s='http://purl.oclc.org/dsdl/svrl'><s:fired-rule context='Invoice'/><s:failed-assert id='CII-SR-465' flag='warning' location='/Invoice'><s:text>Contact cardinality</s:text></s:failed-assert></s:schematron-output>";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(svrl));
        var result = SaxonInvoiceRulesRunner.ReadSvrl(stream, new Dictionary<string, InvoiceDiagnosticSeverity> { ["CII-SR-465"] = InvoiceDiagnosticSeverity.Error });
        Assert.Equal(InvoiceDiagnosticSeverity.Error, Assert.Single(result).Severity);
    }

    [Fact]
    public void CorruptAuthorityBundleIsRejectedBeforeZipParsing() {
        string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-invoice-test-" + Guid.NewGuid().ToString("N"));
        try {
            File.WriteAllBytes(path, new byte[] { 1, 2, 3 });
            Assert.Throws<InvalidDataException>(() => InvoiceRuleBundle.Load(path));
        } finally { File.Delete(path); }
    }
}
