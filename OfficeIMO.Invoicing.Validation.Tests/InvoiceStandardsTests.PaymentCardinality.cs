using System.Text;
using System.Xml.Linq;
using OfficeIMO.Invoicing.Tests;

namespace OfficeIMO.Invoicing.Validation.Tests;

public partial class InvoiceStandardsTests {
    [InvoiceStandardsTheory]
    [InlineData(InvoiceSyntax.Cii, InvoiceProfile.En16931)]
    [InlineData(InvoiceSyntax.Ubl, InvoiceProfile.En16931)]
    [InlineData(InvoiceSyntax.Cii, InvoiceProfile.XRechnung)]
    [InlineData(InvoiceSyntax.Ubl, InvoiceProfile.XRechnung)]
    [InlineData(InvoiceSyntax.Ubl, InvoiceProfile.PeppolBis)]
    public async Task PartyIdentifierCardinalityFollowsTheSelectedSyntax(InvoiceSyntax syntax, InvoiceProfile profile) {
        Invoice invoice = PaymentProfileFixture();
        invoice.Seller.Identifiers.Add(new InvoiceIdentifier("seller-1"));
        invoice.Seller.Identifiers.Add(new InvoiceIdentifier("seller-2"));
        invoice.Buyer.Identifiers.Add(new InvoiceIdentifier("buyer-1"));
        if (syntax == InvoiceSyntax.Cii) invoice.Buyer.Identifiers.Add(new InvoiceIdentifier("buyer-2"));
        var options = new InvoiceXmlOptions(syntax, profile);
        var validator = new InvoiceValidator(Bundle(), Runner());
        byte[] xml = InvoiceSerializer.Write(invoice, options);
        InvoiceValidationReport report = await validator.ValidateAsync(xml, PaymentProfileRelease(profile));
        Assert.True(report.IsValid, Report(report));
        InvoiceReadResult read = InvoiceParser.Read(xml);
        Assert.True(read.HasCompleteMapping);
        Assert.Equal(2, read.Invoice.Seller.Identifiers.Count);
        Assert.Equal(syntax == InvoiceSyntax.Cii ? 2 : 1, read.Invoice.Buyer.Identifiers.Count);
        if (syntax == InvoiceSyntax.Ubl) {
            XNamespace cac = "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2";
            XNamespace cbc = "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2";
            XDocument document = XDocument.Parse(Encoding.UTF8.GetString(xml));
            document.Descendants(cac + "AccountingCustomerParty").Single().Element(cac + "Party")!
                .Element(cac + "PartyIdentification")!.AddAfterSelf(new XElement(cac + "PartyIdentification", new XElement(cbc + "ID", "buyer-2")));
            xml = Encoding.UTF8.GetBytes(document.ToString());
            report = await validator.ValidateAsync(xml, PaymentProfileRelease(profile));
            Assert.False(report.IsValid);
            Assert.Contains(report.Diagnostics, d => d.Code == "UBL-SR-16");
        }
        InvoiceConversionResult conversion = InvoiceConverter.Convert(xml, new InvoiceXmlOptions(InvoiceSyntax.Ubl, profile));
        Assert.False(conversion.Succeeded);
        Assert.Null(conversion.Xml);
        Assert.Contains(conversion.Diagnostics, d => d.Location == "Buyer.Identifiers");
    }

    [InvoiceStandardsTheory]
    [InlineData(InvoiceSyntax.Cii, InvoiceProfile.En16931, true)]
    [InlineData(InvoiceSyntax.Ubl, InvoiceProfile.En16931, true)]
    [InlineData(InvoiceSyntax.Cii, InvoiceProfile.XRechnung, true)]
    [InlineData(InvoiceSyntax.Ubl, InvoiceProfile.XRechnung, true)]
    [InlineData(InvoiceSyntax.Ubl, InvoiceProfile.PeppolBis, true)]
    [InlineData(InvoiceSyntax.Ubl, InvoiceProfile.PeppolBis, false)]
    public async Task DirectDebitRequirementsMatchThePinnedProfile(InvoiceSyntax syntax, InvoiceProfile profile, bool german) {
        Invoice invoice = PaymentProfileFixture();
        if (!german) {
            invoice.Seller.Address!.CountryCode = "FR";
            invoice.Buyer.Address!.CountryCode = "FR";
            invoice.Seller.VatIdentifier = "FR40303265045";
        }
        invoice.Payment!.MeansCode = "59";
        invoice.Payment.Accounts.Clear();
        invoice.Payment.MandateReference = "mandate-1";
        invoice.Payment.CreditorIdentifier = "DE98ZZZ09999999999";
        invoice.Payment.DebitedAccount = "DE89370400440532013000";
        var options = new InvoiceXmlOptions(syntax, profile);
        byte[] xml = InvoiceSerializer.Write(invoice, options);
        var validator = new InvoiceValidator(Bundle(), Runner());
        InvoiceValidationReport report = await validator.ValidateAsync(xml, PaymentProfileRelease(profile));
        Assert.True(report.IsValid, Report(report));
        for (int field = 0; field < 3; field++) {
            XDocument document = XDocument.Parse(Encoding.UTF8.GetString(xml));
            if (field == 0) {
                if (syntax == InvoiceSyntax.Cii) document.Descendants().Single(e => e.Name.LocalName == "DirectDebitMandateID").Remove();
                else document.Descendants().Single(e => e.Name.LocalName == "PaymentMandate").Elements().Single(e => e.Name.LocalName == "ID").Remove();
            } else if (field == 1) {
                if (syntax == InvoiceSyntax.Cii) document.Descendants().Single(e => e.Name.LocalName == "CreditorReferenceID").Remove();
                else document.Descendants().Single(e => (string?)e.Attribute("schemeID") == "SEPA").Parent!.Remove();
            } else document.Descendants().Single(e => e.Name.LocalName == (syntax == InvoiceSyntax.Cii ? "PayerPartyDebtorFinancialAccount" : "PayerFinancialAccount")).Remove();
            byte[] changed = Encoding.UTF8.GetBytes(document.ToString());
            bool required = profile != InvoiceProfile.En16931 && (field == 0 || profile == InvoiceProfile.XRechnung || german);
            report = await validator.ValidateAsync(changed, PaymentProfileRelease(profile));
            Assert.True(report.IsValid == !required, field + ": " + Report(report));
            InvoiceReadResult read = InvoiceParser.Read(changed);
            Assert.True(read.HasCompleteMapping, string.Join("; ", read.UnmappedData.Select(d => d.Message)));
            InvoiceConversionResult conversion = InvoiceConverter.Convert(changed, options);
            Assert.Equal(!required, conversion.Succeeded);
            if (required) {
                string path = field == 0 ? "Payment.MandateReference" : field == 1 ? "Payment.CreditorIdentifier" : "Payment.DebitedAccount";
                Assert.Contains(conversion.Diagnostics, d => d.Location == path);
                Assert.Throws<InvalidDataException>(() => read.Write(options));
            }
        }
    }

    private static Invoice PaymentProfileFixture() {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Seller.ElectronicAddress = new InvoiceIdentifier("1234567890128", "0088");
        invoice.Buyer.ElectronicAddress = new InvoiceIdentifier("1234567890135", "0088");
        return invoice;
    }

    private static InvoiceRulesRelease PaymentProfileRelease(InvoiceProfile profile) => profile switch {
        InvoiceProfile.En16931 => InvoiceRulesRelease.En16931_1_3_16,
        InvoiceProfile.XRechnung => InvoiceRulesRelease.XRechnung_3_0_2_2026_08_31,
        _ => InvoiceRulesRelease.PeppolBis_3_0_21
    };
}
