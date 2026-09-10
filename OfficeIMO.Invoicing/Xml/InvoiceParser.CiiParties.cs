using System.Xml.Linq;

namespace OfficeIMO.Invoicing;

public static partial class InvoiceParser {
    private static InvoiceParty CiiParty(InvoiceXmlReadContext c, XElement? element) {
        var party = new InvoiceParty { Name = c.Required(element, Ram + "Name"), LegalInformation = c.Text(element, Ram + "Description"),
            Address = CiiAddress(c, c.Child(element, Ram + "PostalTradeAddress")),
            ElectronicAddress = c.Identifier(c.Child(c.Child(element, Ram + "URIUniversalCommunication"), Ram + "URIID")) };
        foreach (XElement identifier in c.Children(element, Ram + "ID")) party.Identifiers.Add(c.Identifier(identifier)!);
        foreach (XElement identifier in c.Children(element, Ram + "GlobalID")) party.Identifiers.Add(c.Identifier(identifier)!);
        XElement? legal = c.Child(element, Ram + "SpecifiedLegalOrganization");
        party.LegalRegistration = c.Identifier(c.Child(legal, Ram + "ID")); party.TradingName = c.Text(legal, Ram + "TradingBusinessName");
        XElement? contact = c.Child(element, Ram + "DefinedTradeContact");
        if (contact != null) party.Contact = new InvoiceContact { Name = c.Text(contact, Ram + "PersonName"),
            Telephone = c.Text(c.Child(contact, Ram + "TelephoneUniversalCommunication"), Ram + "CompleteNumber"), Email = c.Text(c.Child(contact, Ram + "EmailURIUniversalCommunication"), Ram + "URIID") };
        foreach (XElement tax in c.Children(element, Ram + "SpecifiedTaxRegistration")) {
            XElement? identifier = c.Child(tax, Ram + "ID");
            string? value = c.Value(identifier), scheme = c.Attribute(identifier, "schemeID");
            if (scheme == "VA") {
                if (party.VatIdentifier != null) c.Loss(tax, "Multiple VAT identifiers are outside the supported mapping.");
                party.VatIdentifier = value;
            } else if (scheme == "FC") {
                if (party.TaxRegistration != null) c.Loss(tax, "Multiple non-VAT tax registrations are outside the supported mapping.");
                party.TaxRegistration = value;
            } else c.Loss(tax, "Unsupported tax registration scheme.");
        }
        return party;
    }
    private static InvoiceAddress CiiAddress(InvoiceXmlReadContext c, XElement? element) => new InvoiceAddress {
        Line1 = c.Text(element, Ram + "LineOne"), Line2 = c.Text(element, Ram + "LineTwo"), Line3 = c.Text(element, Ram + "LineThree"),
        City = c.Text(element, Ram + "CityName"), PostCode = c.Text(element, Ram + "PostcodeCode"), CountryCode = c.Required(element, Ram + "CountryID"), Subdivision = c.Text(element, Ram + "CountrySubDivisionName")
    };

    private static InvoicePayment? CiiPayment(InvoiceXmlReadContext c, XElement? settlement) {
        string? reference = c.Text(settlement, Ram + "PaymentReference"), creditor = c.Text(settlement, Ram + "CreditorReferenceID");
        InvoicePayment? result = null;
        foreach (XElement element in c.Children(settlement, Ram + "SpecifiedTradeSettlementPaymentMeans")) {
            string code = c.Required(element, Ram + "TypeCode");
            string? text = c.Text(element, Ram + "Information");
            if (result == null) result = new InvoicePayment { MeansCode = code, MeansText = text, Reference = reference, CreditorIdentifier = creditor };
            else { Agree(c, element, result.MeansCode, code, "payment means"); if (result.MeansText != null && text != null) Agree(c, element, result.MeansText, text, "payment descriptions"); result.MeansText = result.MeansText ?? text; }
            XElement? account = c.Child(element, Ram + "PayeePartyCreditorFinancialAccount");
            XElement? institution = c.Child(element, Ram + "PayeeSpecifiedCreditorFinancialInstitution");
            if (account != null) {
                string? iban = c.Text(account, Ram + "IBANID"), local = c.Text(account, Ram + "ProprietaryID");
                if (iban != null && local != null) c.Loss(account, "Account declares both IBAN and proprietary identifiers.");
                result.Accounts.Add(new InvoiceBankAccount { Identifier = iban ?? local ?? string.Empty, IsIban = iban != null,
                    Name = c.Text(account, Ram + "AccountName"), ProviderIdentifier = c.Text(institution, Ram + "BICID") });
            } else if (institution != null) c.Loss(institution, "A financial institution without an account is outside the supported mapping.");
            XElement? card = c.Child(element, Ram + "ApplicableTradeSettlementFinancialCard");
            string? cardNumber = c.Text(card, Ram + "ID"), cardHolder = c.Text(card, Ram + "CardholderName");
            string? debit = c.Text(c.Child(element, Ram + "PayerPartyDebtorFinancialAccount"), Ram + "IBANID");
            if (result.CardNumber != null && cardNumber != null) Agree(c, element, result.CardNumber, cardNumber, "card numbers");
            if (result.CardHolder != null && cardHolder != null) Agree(c, element, result.CardHolder, cardHolder, "card holders");
            if (result.DebitedAccount != null && debit != null) Agree(c, element, result.DebitedAccount, debit, "debited accounts");
            result.CardNumber = result.CardNumber ?? cardNumber; result.CardHolder = result.CardHolder ?? cardHolder; result.DebitedAccount = result.DebitedAccount ?? debit;
        }
        return result ?? (reference == null && creditor == null ? null : new InvoicePayment { Reference = reference, CreditorIdentifier = creditor });
    }
}
