using System.Xml.Linq;

namespace OfficeIMO.Invoicing;

public static partial class InvoiceParser {
    private static InvoiceParty CiiParty(InvoiceXmlReadContext c, XElement? element) {
        var party = new InvoiceParty { Name = c.Required(element, Ram + "Name"), LegalInformation = c.Text(element, Ram + "Description"),
            Address = CiiAddress(c, c.Child(element, Ram + "PostalTradeAddress")),
            ElectronicAddress = c.Identifier(c.Child(c.Child(element, Ram + "URIUniversalCommunication"), Ram + "URIID")) };
        foreach (XElement identifier in c.Children(element, Ram + "ID")) c.AddTo(party.Identifiers, c.Identifier(identifier)!);
        foreach (XElement identifier in c.Children(element, Ram + "GlobalID")) c.AddTo(party.Identifiers, c.Identifier(identifier)!);
        XElement? legal = c.Child(element, Ram + "SpecifiedLegalOrganization");
        party.LegalRegistration = c.Identifier(c.Child(legal, Ram + "ID")); party.TradingName = c.Text(legal, Ram + "TradingBusinessName");
        XElement? contact = c.Child(element, Ram + "DefinedTradeContact");
        if (contact != null) party.Contact = new InvoiceContact { Name = c.Text(contact, Ram + "PersonName"),
            Telephone = c.Text(c.Child(contact, Ram + "TelephoneUniversalCommunication"), Ram + "CompleteNumber"), Email = c.Text(c.Child(contact, Ram + "EmailURIUniversalCommunication"), Ram + "URIID") };
        foreach (XElement tax in c.Children(element, Ram + "SpecifiedTaxRegistration")) {
            XElement? identifier = c.Child(tax, Ram + "ID");
            string? value = c.Value(identifier), scheme = c.Attribute(identifier, "schemeID");
            if (value != null && scheme != null) c.AddTo(party.TaxRegistrations, scheme switch {
                "VA" => new InvoiceTaxRegistration(value, InvoiceTaxRegistration.VatScheme, InvoiceTaxRegistrationKind.Vat),
                "FC" => new InvoiceTaxRegistration(value, InvoiceTaxRegistration.TaxScheme, InvoiceTaxRegistrationKind.Fiscal),
                _ => new InvoiceTaxRegistration(value, scheme, InvoiceTaxRegistrationKind.Other)
            });
            else c.Loss(tax, "A tax registration requires both an identifier and a scheme to preserve its meaning.");
        }
        return party;
    }
    private static InvoiceAddress CiiAddress(InvoiceXmlReadContext c, XElement? element) => new InvoiceAddress {
        Line1 = c.Text(element, Ram + "LineOne"), Line2 = c.Text(element, Ram + "LineTwo"), Line3 = c.Text(element, Ram + "LineThree"),
        City = c.Text(element, Ram + "CityName"), PostCode = c.Text(element, Ram + "PostcodeCode"), CountryCode = c.Required(element, Ram + "CountryID"), Subdivision = c.Text(element, Ram + "CountrySubDivisionName")
    };

    private static IEnumerable<InvoicePayment> CiiPayments(InvoiceXmlReadContext c, XElement? settlement) {
        string? reference = c.Text(settlement, Ram + "PaymentReference"), creditor = c.Text(settlement, Ram + "CreditorReferenceID");
        bool found = false;
        foreach (XElement element in c.Children(settlement, Ram + "SpecifiedTradeSettlementPaymentMeans")) {
            found = true;
            var result = new InvoicePayment {
                MeansCode = c.Required(element, Ram + "TypeCode"), MeansText = c.Text(element, Ram + "Information"),
                Reference = reference, CreditorIdentifier = creditor
            };
            XElement? account = c.Child(element, Ram + "PayeePartyCreditorFinancialAccount");
            XElement? institution = c.Child(element, Ram + "PayeeSpecifiedCreditorFinancialInstitution");
            if (account != null) {
                string? iban = c.Text(account, Ram + "IBANID"), local = c.Text(account, Ram + "ProprietaryID");
                if (iban != null && local != null) c.Loss(account, "Account declares both IBAN and proprietary identifiers.");
                result.Account = new InvoiceBankAccount { Identifier = iban ?? local ?? string.Empty, IsIban = iban != null,
                    Name = c.Text(account, Ram + "AccountName"), ProviderIdentifier = c.Text(institution, Ram + "BICID") };
            } else if (institution != null) c.Loss(institution, "A financial institution without an account is outside the supported mapping.");
            XElement? card = c.Child(element, Ram + "ApplicableTradeSettlementFinancialCard");
            result.CardNumber = c.Text(card, Ram + "ID");
            result.CardHolder = c.Text(card, Ram + "CardholderName");
            XElement? debtorAccount = c.Child(element, Ram + "PayerPartyDebtorFinancialAccount");
            result.DebitedAccount = c.Text(debtorAccount, Ram + "IBANID");
            if (result.DebitedAccount != null && !InvoiceBankAccountIdentity.IsValidIban(result.DebitedAccount))
                c.Loss(debtorAccount!, "The source debtor account is explicitly identified as an IBAN but does not have a registered country format and valid checksum.");
            yield return result;
        }
        if (!found && (reference != null || creditor != null)) yield return new InvoicePayment { Reference = reference, CreditorIdentifier = creditor };
    }
}
