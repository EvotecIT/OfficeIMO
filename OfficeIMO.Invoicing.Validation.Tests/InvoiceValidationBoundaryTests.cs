using System.Text;

namespace OfficeIMO.Invoicing.Validation.Tests;

public class InvoiceValidationBoundaryTests {
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
