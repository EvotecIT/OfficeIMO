using System.Xml.Linq;

namespace OfficeIMO.Invoicing;

public static partial class InvoiceParser {
    private static InvoiceParty UblParty(InvoiceXmlReadContext c, XElement? element, out string? creditor) {
        XElement? legal = c.Child(element, Cac + "PartyLegalEntity");
        var party = new InvoiceParty { Name = c.Required(legal, Cbc + "RegistrationName"), LegalRegistration = c.Identifier(c.Child(legal, Cbc + "CompanyID")),
            LegalInformation = c.Text(legal, Cbc + "CompanyLegalForm"), TradingName = c.Text(c.Child(element, Cac + "PartyName"), Cbc + "Name"),
            ElectronicAddress = c.Identifier(c.Child(element, Cbc + "EndpointID")), Address = UblAddress(c, c.Child(element, Cac + "PostalAddress")) };
        creditor = null;
        foreach (XElement identification in c.Children(element, Cac + "PartyIdentification")) {
            InvoiceIdentifier identifier = c.Identifier(c.Child(identification, Cbc + "ID")) ?? new InvoiceIdentifier(string.Empty);
            if (identifier.SchemeId == "SEPA") {
                if (creditor != null) c.Loss(identification, "Multiple SEPA creditor identifiers are outside the supported mapping.");
                creditor = identifier.Value;
            } else c.AddTo(party.Identifiers, identifier);
        }
        XElement? contact = c.Child(element, Cac + "Contact");
        if (contact != null) party.Contact = new InvoiceContact { Name = c.Text(contact, Cbc + "Name"), Telephone = c.Text(contact, Cbc + "Telephone"), Email = c.Text(contact, Cbc + "ElectronicMail") };
        UblTaxRegistrations(c, element, party);
        return party;
    }
    private static InvoiceAddress UblAddress(InvoiceXmlReadContext c, XElement? element) => new InvoiceAddress {
        Line1 = c.Text(element, Cbc + "StreetName"), Line2 = c.Text(element, Cbc + "AdditionalStreetName"), Line3 = c.Text(c.Child(element, Cac + "AddressLine"), Cbc + "Line"),
        City = c.Text(element, Cbc + "CityName"), PostCode = c.Text(element, Cbc + "PostalZone"), Subdivision = c.Text(element, Cbc + "CountrySubentity"),
        CountryCode = c.Required(c.Child(element, Cac + "Country"), Cbc + "IdentificationCode")
    };
    private static void UblTaxRegistrations(InvoiceXmlReadContext c, XElement? element, InvoiceParty party) {
        foreach (XElement tax in c.Children(element, Cac + "PartyTaxScheme")) {
            string? scheme = c.Text(c.Child(tax, Cac + "TaxScheme"), Cbc + "ID"), identifier = c.Text(tax, Cbc + "CompanyID");
            if (scheme == "VAT") {
                if (party.VatIdentifier != null) c.Loss(tax, "Multiple VAT identifiers are outside the supported mapping.");
                party.VatIdentifier = identifier;
            } else {
                if (scheme != "TAX") c.Loss(tax, "The non-VAT tax scheme cannot be preserved by the supported TAX/FC mapping.");
                if (party.TaxRegistration != null) c.Loss(tax, "Multiple non-VAT registrations are outside the supported mapping.");
                party.TaxRegistration = identifier;
            }
        }
    }
    private static InvoicePayment? UblPayment(InvoiceXmlReadContext c, XElement root, string? creditor) {
        InvoicePayment? result = null;
        foreach (XElement element in c.Children(root, Cac + "PaymentMeans")) {
            XElement? code = c.Child(element, Cbc + "PaymentMeansCode");
            string means = c.Value(code) ?? string.Empty;
            string? text = c.Attribute(code, "name"), reference = c.Text(element, Cbc + "PaymentID");
            if (result == null) result = new InvoicePayment { MeansCode = means, MeansText = text, Reference = reference, CreditorIdentifier = creditor };
            else {
                Agree(c, element, result.MeansCode, means, "payment means");
                if (result.MeansText != null && text != null) Agree(c, element, result.MeansText, text, "payment descriptions");
                if (result.Reference != null && reference != null) Agree(c, element, result.Reference, reference, "payment references");
                result.MeansText = result.MeansText ?? text;
                result.Reference = result.Reference ?? reference;
            }
            XElement? account = c.Child(element, Cac + "PayeeFinancialAccount");
            if (account != null) {
                string identifier = c.Required(account, Cbc + "ID");
                c.AddTo(result.Accounts, new InvoiceBankAccount { Identifier = identifier, IsIban = InvoiceBankAccountIdentity.IsValidIban(identifier),
                    Name = c.Text(account, Cbc + "Name"), ProviderIdentifier = c.Text(c.Child(account, Cac + "FinancialInstitutionBranch"), Cbc + "ID") });
            }
            XElement? card = c.Child(element, Cac + "CardAccount"), mandate = c.Child(element, Cac + "PaymentMandate");
            c.Expected(c.Child(card, Cbc + "NetworkID"), "NA");
            string? cardNumber = c.Text(card, Cbc + "PrimaryAccountNumberID"), cardHolder = c.Text(card, Cbc + "HolderName");
            string? mandateReference = c.Text(mandate, Cbc + "ID"), debit = c.Text(c.Child(mandate, Cac + "PayerFinancialAccount"), Cbc + "ID");
            if (result.CardNumber != null && cardNumber != null) Agree(c, element, result.CardNumber, cardNumber, "card numbers");
            if (result.CardHolder != null && cardHolder != null) Agree(c, element, result.CardHolder, cardHolder, "card holders");
            if (result.MandateReference != null && mandateReference != null) Agree(c, element, result.MandateReference, mandateReference, "mandates");
            if (result.DebitedAccount != null && debit != null) Agree(c, element, result.DebitedAccount, debit, "debited accounts");
            result.CardNumber = result.CardNumber ?? cardNumber; result.CardHolder = result.CardHolder ?? cardHolder;
            result.MandateReference = result.MandateReference ?? mandateReference; result.DebitedAccount = result.DebitedAccount ?? debit;
        }
        return result ?? (creditor == null ? null : new InvoicePayment { CreditorIdentifier = creditor });
    }
}
