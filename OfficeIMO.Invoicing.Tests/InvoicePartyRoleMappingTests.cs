using System.Text;
using System.Xml.Linq;

namespace OfficeIMO.Invoicing.Tests;

public class InvoicePartyRoleMappingTests {
    [Theory]
    [InlineData(InvoiceSyntax.Cii, "Buyer", "LegalInformation")]
    [InlineData(InvoiceSyntax.Cii, "Buyer", "TaxRegistration")]
    [InlineData(InvoiceSyntax.Ubl, "Buyer", "LegalInformation")]
    [InlineData(InvoiceSyntax.Ubl, "Buyer", "TaxRegistration")]
    [InlineData(InvoiceSyntax.Cii, "Payee", "Contact")]
    [InlineData(InvoiceSyntax.Cii, "Payee", "VatIdentifier")]
    [InlineData(InvoiceSyntax.Cii, "TaxRepresentative", "LegalRegistration")]
    [InlineData(InvoiceSyntax.Cii, "TaxRepresentative", "TradingName")]
    [InlineData(InvoiceSyntax.Ubl, "TaxRepresentative", "TaxRegistration")]
    public void RoleInvalidSourceFieldsAreReportedAndCanBeExplicitlyDiscarded(InvoiceSyntax syntax, string role, string field) {
        var options = new InvoiceXmlOptions(syntax);
        XDocument document = XDocument.Parse(Encoding.UTF8.GetString(InvoiceSerializer.Write(InvoiceFixture.Rich(), options)));
        XNamespace ram = "urn:un:unece:uncefact:data:standard:ReusableAggregateBusinessInformationEntity:100";
        XNamespace cac = "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2";
        XNamespace cbc = "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2";
        if (syntax == InvoiceSyntax.Cii) {
            string name = role == "Buyer" ? "BuyerTradeParty" : role == "Payee" ? "PayeeTradeParty" : "SellerTaxRepresentativeTradeParty";
            XElement party = document.Descendants(ram + name).Single();
            if (field == "LegalInformation") party.Add(new XElement(ram + "Description", "unsupported"));
            else if (field == "TaxRegistration" || field == "VatIdentifier")
                party.Add(new XElement(ram + "SpecifiedTaxRegistration", new XElement(ram + "ID", new XAttribute("schemeID", field == "TaxRegistration" ? "FC" : "VA"), "unsupported")));
            else if (field == "Contact") party.Add(new XElement(ram + "DefinedTradeContact", new XElement(ram + "PersonName", "unsupported")));
            else party.Add(new XElement(ram + "SpecifiedLegalOrganization", new XElement(ram + (field == "LegalRegistration" ? "ID" : "TradingBusinessName"), "unsupported")));
        } else {
            XElement party = role == "Buyer" ? document.Descendants(cac + "AccountingCustomerParty").Single().Element(cac + "Party")! : document.Descendants(cac + "TaxRepresentativeParty").Single();
            if (field == "LegalInformation") party.Element(cac + "PartyLegalEntity")!.Add(new XElement(cbc + "CompanyLegalForm", "unsupported"));
            else party.Add(new XElement(cac + "PartyTaxScheme", new XElement(cbc + "CompanyID", "unsupported"), new XElement(cac + "TaxScheme", new XElement(cbc + "ID", "TAX"))));
        }
        byte[] xml = Encoding.UTF8.GetBytes(document.ToString());
        InvoiceReadResult read = InvoiceParser.Read(xml);
        Assert.False(read.HasCompleteMapping);
        Assert.Equal(xml, read.GetOriginalBytes());
        Assert.Throws<InvalidDataException>(() => read.Write());
        Assert.False(InvoiceConverter.Convert(xml, options).Succeeded);
        InvoiceReadResult rewritten = InvoiceParser.Read(read.Write(options, allowUnmappedDataLoss: true));
        Assert.True(rewritten.HasCompleteMapping);
        Assert.DoesNotContain("unsupported", Encoding.UTF8.GetString(rewritten.GetOriginalBytes()), StringComparison.Ordinal);
    }
}
