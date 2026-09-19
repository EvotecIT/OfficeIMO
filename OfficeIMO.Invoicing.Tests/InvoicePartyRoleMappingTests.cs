using System.Text;
using System.Xml.Linq;

namespace OfficeIMO.Invoicing.Tests;

public class InvoicePartyRoleMappingTests {
    [Theory]
    [InlineData(InvoiceSyntax.Cii, "Buyer", "LegalInformation", "Buyer.LegalInformation")]
    [InlineData(InvoiceSyntax.Cii, "Buyer", "TaxRegistration", "Buyer.TaxRegistrations[1]")]
    [InlineData(InvoiceSyntax.Ubl, "Buyer", "LegalInformation", "Buyer.LegalInformation")]
    [InlineData(InvoiceSyntax.Ubl, "Buyer", "TaxRegistration", "Buyer.TaxRegistrations[1]")]
    [InlineData(InvoiceSyntax.Cii, "Payee", "Contact", "Payee.Contact")]
    [InlineData(InvoiceSyntax.Cii, "Payee", "VatIdentifier", "Payee.TaxRegistrations")]
    [InlineData(InvoiceSyntax.Cii, "TaxRepresentative", "LegalRegistration", "TaxRepresentative.LegalRegistration")]
    [InlineData(InvoiceSyntax.Cii, "TaxRepresentative", "TradingName", "TaxRepresentative.TradingName")]
    [InlineData(InvoiceSyntax.Ubl, "TaxRepresentative", "TaxRegistration", "TaxRepresentative.TaxRegistrations[1]")]
    public void RoleSourceFieldsArePreservedAndUnsupportedTargetsAreExplicit(InvoiceSyntax syntax, string role, string field, string? unsupportedLocation) {
        var options = InvoiceTestContracts.En16931(syntax);
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
        Assert.True(read.HasCompleteMapping);
        Assert.Equal(xml, read.GetOriginalBytes());
        Assert.True(ContainsImportedValue(read.Invoice, role, field));
        IReadOnlyList<InvoiceDiagnostic> diagnostics = InvoiceSerializer.InspectTarget(read.Invoice, options);
        if (unsupportedLocation == null) {
            Assert.Empty(diagnostics);
            Assert.Contains("unsupported", Encoding.UTF8.GetString(read.Write(options)), StringComparison.Ordinal);
        } else {
            Assert.Contains(diagnostics, diagnostic => diagnostic.Location == unsupportedLocation);
            Assert.Throws<InvalidDataException>(() => read.Write(options));
            Assert.False(InvoiceConverter.Convert(xml, options).Succeeded);
        }
    }

    private static bool ContainsImportedValue(Invoice invoice, string role, string field) {
        InvoiceParty party = role == "Buyer" ? invoice.Buyer : role == "Payee" ? invoice.Payee! : invoice.TaxRepresentative!;
        return field switch {
            "LegalInformation" => party.LegalInformation == "unsupported",
            "TaxRegistration" => party.TaxRegistrations.Any(registration => registration.Identifier == "unsupported"),
            "VatIdentifier" => party.TaxRegistrations.Any(registration => registration.Identifier == "unsupported"),
            "Contact" => party.Contact?.Name == "unsupported",
            "LegalRegistration" => party.LegalRegistration?.Value == "unsupported",
            "TradingName" => party.TradingName == "unsupported",
            _ => false
        };
    }
}
