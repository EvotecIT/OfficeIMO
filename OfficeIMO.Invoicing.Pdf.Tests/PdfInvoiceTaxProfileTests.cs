using System.Text;
using System.Xml.Linq;
using OfficeIMO.Invoicing.Tests;
using OfficeIMO.Pdf;
using OfficeIMO.Tests.Pdf;
using Xunit;

namespace OfficeIMO.Invoicing.Pdf.Tests;

public sealed class PdfInvoiceTaxProfileTests {
    private static readonly XNamespace Ram = "urn:un:unece:uncefact:data:standard:ReusableAggregateBusinessInformationEntity:100";

    [Theory]
    [InlineData(InvoiceProfile.BasicWithoutLines)]
    [InlineData(InvoiceProfile.Basic)]
    [InlineData(InvoiceProfile.En16931)]
    public void StrictInvoiceAcceptsOutsideScopeTaxWithoutRate(InvoiceProfile profile) {
        PdfInvoiceDocument snapshot = PdfInvoiceDocument.Create(InvoiceFixture.WithTaxCategory("O"), Contract(profile));
        byte[] xml = snapshot.ToXmlBytes();
        XDocument parsed = XDocument.Parse(Encoding.UTF8.GetString(xml));
        XElement headerTax = Assert.Single(parsed.Descendants(Ram + "ApplicableHeaderTradeSettlement")
            .Elements(Ram + "ApplicableTradeTax"));
        Assert.Equal("O", Assert.Single(headerTax.Elements(Ram + "CategoryCode")).Value);
        Assert.Empty(headerTax.Elements(Ram + "RateApplicablePercent"));

        byte[] pdf = snapshot.ToPdfBytes(FontOptions().UsePdfA(PdfComplianceProfile.PdfA3B));
        AssertInvoicePair(profile, xml, pdf);
        Export("outside-scope-" + profile.ToString().ToLowerInvariant(), xml, pdf);
    }

    [Theory]
    [InlineData(InvoiceProfile.BasicWithoutLines, false)]
    [InlineData(InvoiceProfile.BasicWithoutLines, true)]
    [InlineData(InvoiceProfile.Basic, false)]
    [InlineData(InvoiceProfile.Basic, true)]
    [InlineData(InvoiceProfile.En16931, false)]
    [InlineData(InvoiceProfile.En16931, true)]
    public void StrictProfileAcceptsInvoiceAndAccountingCurrencyTaxTotalsInEitherOrder(
        InvoiceProfile profile, bool accountingCurrencyFirst) {
        byte[] xml = Encoding.UTF8.GetBytes(DualCurrencyXml(profile, accountingCurrencyFirst).ToString());
        byte[] pdf = RawDocument(xml).ToBytes();
        AssertInvoicePair(profile, xml, pdf);
        Export("dual-currency-" + profile.ToString().ToLowerInvariant() +
            (accountingCurrencyFirst ? "-usd-first" : "-eur-first"), xml, pdf);
    }

    [Theory]
    [InlineData(InvoiceProfile.BasicWithoutLines, "EUR")]
    [InlineData(InvoiceProfile.BasicWithoutLines, "USD")]
    [InlineData(InvoiceProfile.Basic, "EUR")]
    [InlineData(InvoiceProfile.Basic, "USD")]
    [InlineData(InvoiceProfile.En16931, "EUR")]
    [InlineData(InvoiceProfile.En16931, "USD")]
    public void StrictProfileRejectsDuplicateTaxTotalCurrency(InvoiceProfile profile, string currency) {
        XDocument xml = DualCurrencyXml(profile);
        XElement totals = Assert.Single(xml.Descendants(Ram + "SpecifiedTradeSettlementHeaderMonetarySummation"));
        XElement tax = Assert.Single(totals.Elements(Ram + "TaxTotalAmount"),
            element => (string?)element.Attribute("currencyID") == currency);
        tax.AddAfterSelf(new XElement(tax));
        Assert.Throws<InvalidOperationException>(() => RawDocument(Encoding.UTF8.GetBytes(xml.ToString())).ToBytes());
    }

    private static InvoiceXmlOptions Contract(InvoiceProfile profile) => new(
        InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2, InvoiceSyntax.Cii, profile,
        profile == InvoiceProfile.En16931 ? InvoiceProjectionPolicy.RejectDataLoss
            : InvoiceProjectionPolicy.AllowProfileDefinedDataLoss);

    private static XDocument DualCurrencyXml(InvoiceProfile profile, bool accountingCurrencyFirst = false) {
        XDocument xml = XDocument.Parse(Encoding.UTF8.GetString(
            PdfInvoiceDocument.Create(InvoiceFixture.Create(), Contract(profile)).ToXmlBytes()));
        XElement settlement = Assert.Single(xml.Descendants(Ram + "ApplicableHeaderTradeSettlement"));
        XElement invoiceCurrency = Assert.Single(settlement.Elements(Ram + "InvoiceCurrencyCode"));
        Assert.Equal("EUR", invoiceCurrency.Value);
        invoiceCurrency.AddBeforeSelf(new XElement(Ram + "TaxCurrencyCode", "USD"));
        XElement totals = Assert.Single(settlement.Elements(Ram + "SpecifiedTradeSettlementHeaderMonetarySummation"));
        XElement invoiceTax = Assert.Single(totals.Elements(Ram + "TaxTotalAmount"));
        Assert.Equal("EUR", (string?)invoiceTax.Attribute("currencyID"));
        var accountingTax = new XElement(Ram + "TaxTotalAmount", new XAttribute("currencyID", "USD"), "22.00");
        if (accountingCurrencyFirst) invoiceTax.AddBeforeSelf(accountingTax);
        else invoiceTax.AddAfterSelf(accountingTax);
        return xml;
    }

    private static PdfCreateOptions FontOptions() {
        string? path = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(path);
        byte[] font = File.ReadAllBytes(path!);
        return new PdfCreateOptions { IncludeStandardFontToUnicodeMaps = true }
            .EmbedStandardFont(PdfStandardFont.Helvetica, font, "OfficeIMO Source Serif")
            .EmbedStandardFont(PdfStandardFont.HelveticaBold, font, "OfficeIMO Source Serif");
    }

    private static PdfDocument RawDocument(byte[] xml) {
        PdfCreateOptions options = FontOptions()
            .UseFacturX(xml, relationship: PdfAssociatedFileRelationship.Alternative,
                textFallbacks: PdfTextFallbackFeatures.None)
            .UsePdfA(PdfComplianceProfile.PdfA3B)
            .SetEmbeddedFileModificationDate("factur-x.xml", DateTimeOffset.UtcNow);
        PdfDocument document = PdfDocument.Create(options);
        document.Content.Paragraph(paragraph => paragraph.Text("Invoice tax profile regression"));
        return document;
    }

    private static void AssertInvoicePair(InvoiceProfile profile, byte[] xml, byte[] pdf) {
        PdfDocument saved = PdfDocument.Load(new System.IO.MemoryStream(pdf));
        PdfXmpMetadataInfo xmp = Assert.IsType<PdfXmpMetadataInfo>(saved.Inspect().XmpMetadata);
        Assert.Equal(3, xmp.PdfAPart);
        Assert.Equal("B", xmp.PdfAConformance);
        Assert.Equal(InvoiceProfiles.GetXmpConformanceLevel(profile), xmp.ElectronicInvoiceConformanceLevel);
        Assert.Equal(profile, InvoiceProfileDeclaration.Read(xml).Profile);
        Assert.Equal(xml, Assert.Single(saved.Attachments.Extract()).Bytes);
        Assert.Empty(saved.AssessCompliance(PdfComplianceProfile.PdfA3B).MissingRequirements);
        Assert.Empty(saved.AssessCompliance(PdfComplianceProfile.FacturX).MissingRequirements);
        Assert.True(saved.AssessComplianceProof(PdfComplianceProfile.FacturX).IsInternallyReady);
    }

    private static void Export(string name, byte[] xml, byte[] pdf) {
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_PDF_COMPLIANCE_PROOF_OUTPUT");
        if (string.IsNullOrWhiteSpace(output)) return;
        string directory = Path.Combine(output, "invoice-tax-regressions");
        Directory.CreateDirectory(directory);
        File.WriteAllBytes(Path.Combine(directory, name + ".xml"), xml);
        File.WriteAllBytes(Path.Combine(directory, name + ".pdf"), pdf);
    }
}
