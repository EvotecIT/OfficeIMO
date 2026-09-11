using System.Text;
using System.Xml.Linq;

namespace OfficeIMO.Invoicing.Tests;

public class InvoiceDiagnosticLimitTests {
    [Fact]
    public void TruncationCannotHideAnErrorBehindEarlierWarnings() {
        var buffer = new InvoiceDiagnosticBuffer();
        for (int i = 0; i < 2000; i++) buffer.Add("warning", "warning", "Invoice", InvoiceDiagnosticSeverity.Warning);
        buffer.AddRange(new[] { new InvoiceDiagnostic("late-error", "invalid", "Invoice") });
        InvoiceModelValidationResult result = new InvoiceModelValidationResult(buffer.ToList(), null);
        Assert.False(result.IsValid);
        AssertBounded(result.Diagnostics);
        Assert.Throws<InvalidDataException>(result.ThrowIfInvalid);
    }

    [Fact]
    public void ManyMalformedModelItemsRetainBoundedErrorsAndException() {
        Invoice invoice = InvoiceFixture.Create();
        for (int i = 0; i < 49000; i++) invoice.Lines[0].Classifications.Add(new InvoiceItemClassification());
        InvoiceModelValidationResult result = InvoiceModelValidator.Validate(invoice);
        Assert.False(result.IsValid);
        AssertBounded(result.Diagnostics);
        InvalidDataException error = Assert.Throws<InvalidDataException>(result.ThrowIfInvalid);
        Assert.True(error.Message.Length < 200000);
        AssertBounded(InvoiceSerializer.InspectTarget(invoice, new InvoiceXmlOptions()));
    }

    [Fact]
    public void MappedMalformedXmlCannotAmplifyConversionDiagnostics() {
        XDocument document = XDocument.Parse(Encoding.UTF8.GetString(InvoiceSerializer.Write(InvoiceFixture.Create(), new InvoiceXmlOptions(InvoiceSyntax.Ubl))));
        XNamespace cac = "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2";
        XNamespace cbc = "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2";
        XElement item = document.Descendants(cac + "Item").First();
        for (int i = 0; i < 20000; i++) item.Add(new XElement(cac + "CommodityClassification", new XElement(cbc + "ItemClassificationCode", new XAttribute("listID", ""), "")));
        byte[] xml = Encoding.UTF8.GetBytes(document.ToString());
        Assert.True(xml.Length < 16 * 1024 * 1024);
        InvoiceReadResult read = InvoiceParser.Read(xml);
        Assert.True(read.HasCompleteMapping);
        InvoiceConversionResult result = InvoiceConverter.Convert(xml, new InvoiceXmlOptions(InvoiceSyntax.Cii, InvoiceProfile.XRechnung));
        Assert.False(result.Succeeded);
        Assert.Null(result.Xml);
        AssertBounded(result.Diagnostics);
    }

    [Fact]
    public void TargetDiagnosticsAreBoundedAndNullItemsAreReported() {
        Invoice invoice = InvoiceFixture.Create();
        for (int i = 0; i < 49000; i++) invoice.Payment!.Accounts.Add(new InvoiceBankAccount { Identifier = "DE89370400440532013000", IsIban = false });
        AssertBounded(InvoiceSerializer.InspectTarget(invoice, new InvoiceXmlOptions(InvoiceSyntax.Ubl)));
        invoice.Payment!.Accounts.Clear();
        invoice.Payment.Accounts.Add(null!);
        Assert.Contains(InvoiceSerializer.InspectTarget(invoice, new InvoiceXmlOptions()), d => d.Severity == InvoiceDiagnosticSeverity.Error);
    }

    private static void AssertBounded(IReadOnlyList<InvoiceDiagnostic> diagnostics) {
        Assert.Equal(1000, diagnostics.Count);
        Assert.Contains(diagnostics, d => d.Code == "INV-DIAGNOSTICS-TRUNCATED" && d.Severity == InvoiceDiagnosticSeverity.Error);
    }
}
