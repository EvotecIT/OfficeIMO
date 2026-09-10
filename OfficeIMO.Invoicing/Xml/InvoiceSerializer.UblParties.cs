using System.Xml.Linq;

namespace OfficeIMO.Invoicing;

public static partial class InvoiceSerializer {
    private static XElement UblParty(InvoiceParty party, string? creditorIdentifier = null) => new XElement(Cac + "Party",
        Identifier(Cbc + "EndpointID", party.ElectronicAddress),
        party.Identifiers.Select(identifier => new XElement(Cac + "PartyIdentification", Identifier(Cbc + "ID", identifier))),
        creditorIdentifier == null ? null : new XElement(Cac + "PartyIdentification", new XElement(Cbc + "ID", new XAttribute("schemeID", "SEPA"), creditorIdentifier)),
        party.TradingName == null ? null : new XElement(Cac + "PartyName", Text(Cbc + "Name", party.TradingName)),
        UblAddress("PostalAddress", party.Address), UblTaxRegistration(party.VatIdentifier, "VAT"), UblTaxRegistration(party.TaxRegistration, "TAX"),
        new XElement(Cac + "PartyLegalEntity", Text(Cbc + "RegistrationName", party.Name), Identifier(Cbc + "CompanyID", party.LegalRegistration), Text(Cbc + "CompanyLegalForm", party.LegalInformation)),
        party.Contact == null ? null : new XElement(Cac + "Contact", Text(Cbc + "Name", party.Contact.Name), Text(Cbc + "Telephone", party.Contact.Telephone), Text(Cbc + "ElectronicMail", party.Contact.Email)));
    private static XElement? UblTaxRegistration(string? identifier, string scheme) => identifier == null ? null : new XElement(Cac + "PartyTaxScheme",
        Text(Cbc + "CompanyID", identifier), new XElement(Cac + "TaxScheme", Text(Cbc + "ID", scheme)));
    private static XElement UblAddress(string name, InvoiceAddress address) => new XElement(Cac + name,
        Text(Cbc + "StreetName", address.Line1), Text(Cbc + "AdditionalStreetName", address.Line2), Text(Cbc + "CityName", address.City),
        Text(Cbc + "PostalZone", address.PostCode), Text(Cbc + "CountrySubentity", address.Subdivision),
        address.Line3 == null ? null : new XElement(Cac + "AddressLine", Text(Cbc + "Line", address.Line3)),
        new XElement(Cac + "Country", Text(Cbc + "IdentificationCode", address.CountryCode)));
    private static XElement? UblDelivery(InvoiceDelivery? delivery) => delivery == null ? null : new XElement(Cac + "Delivery",
        UblDate("ActualDeliveryDate", delivery.Date),
        delivery.LocationIdentifier == null && delivery.Address == null ? null : new XElement(Cac + "DeliveryLocation", Identifier(Cbc + "ID", delivery.LocationIdentifier),
            delivery.Address == null ? null : UblAddress("Address", delivery.Address)),
        delivery.Name == null ? null : new XElement(Cac + "DeliveryParty", new XElement(Cac + "PartyName", Text(Cbc + "Name", delivery.Name))));

    private static IEnumerable<XElement> UblPayment(InvoicePayment? payment) {
        if (payment == null) yield break;
        IEnumerable<InvoiceBankAccount?> accounts = payment.Accounts.Count == 0 ? new InvoiceBankAccount?[] { null } : payment.Accounts.Select(account => (InvoiceBankAccount?)account);
        bool first = true;
        foreach (InvoiceBankAccount? account in accounts) {
            yield return new XElement(Cac + "PaymentMeans", new XElement(Cbc + "PaymentMeansCode", !first || payment.MeansText == null ? null : new XAttribute("name", payment.MeansText), payment.MeansCode),
                first ? Text(Cbc + "PaymentID", payment.Reference) : null,
                !first || payment.CardNumber == null ? null : new XElement(Cac + "CardAccount", Text(Cbc + "PrimaryAccountNumberID", payment.CardNumber), new XElement(Cbc + "NetworkID", "NA"), Text(Cbc + "HolderName", payment.CardHolder)),
                account == null ? null : new XElement(Cac + "PayeeFinancialAccount", Text(Cbc + "ID", account.Identifier), Text(Cbc + "Name", account.Name),
                    account.ProviderIdentifier == null ? null : new XElement(Cac + "FinancialInstitutionBranch", Text(Cbc + "ID", account.ProviderIdentifier))),
                !first || payment.MandateReference == null && payment.DebitedAccount == null ? null : new XElement(Cac + "PaymentMandate", Text(Cbc + "ID", payment.MandateReference),
                    payment.DebitedAccount == null ? null : new XElement(Cac + "PayerFinancialAccount", Text(Cbc + "ID", payment.DebitedAccount))));
            first = false;
        }
    }
}
