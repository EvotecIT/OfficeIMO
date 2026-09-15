using System.Xml.Linq;

namespace OfficeIMO.Invoicing;

public static partial class InvoiceSerializer {
    private static XElement CiiParty(string element, InvoiceParty party, bool includeAddress = true) => new XElement(Ram + element,
        party.Identifiers.Where(id => id.SchemeId == null).Select(id => Identifier(Ram + "ID", id)),
        party.Identifiers.Where(id => id.SchemeId != null).Select(id => Identifier(Ram + "GlobalID", id)),
        Text(Ram + "Name", party.Name), Text(Ram + "Description", party.LegalInformation),
        party.LegalRegistration == null && party.TradingName == null ? null : new XElement(Ram + "SpecifiedLegalOrganization", Identifier(Ram + "ID", party.LegalRegistration), Text(Ram + "TradingBusinessName", party.TradingName)),
        party.Contact == null ? null : new XElement(Ram + "DefinedTradeContact", Text(Ram + "PersonName", party.Contact.Name),
            party.Contact.Telephone == null ? null : new XElement(Ram + "TelephoneUniversalCommunication", Text(Ram + "CompleteNumber", party.Contact.Telephone)),
            party.Contact.Email == null ? null : new XElement(Ram + "EmailURIUniversalCommunication", Text(Ram + "URIID", party.Contact.Email))),
        includeAddress ? CiiAddress(party.Address) : null,
        party.ElectronicAddress == null ? null : new XElement(Ram + "URIUniversalCommunication", Identifier(Ram + "URIID", party.ElectronicAddress)),
        party.TaxRegistrations.Select(registration => new XElement(Ram + "SpecifiedTaxRegistration", new XElement(Ram + "ID",
            new XAttribute("schemeID", registration.Kind switch {
                InvoiceTaxRegistrationKind.Vat => "VA",
                InvoiceTaxRegistrationKind.Fiscal => "FC",
                _ => registration.SchemeId
            }), registration.Identifier))));

    private static XElement CiiAddress(InvoiceAddress address) => new XElement(Ram + "PostalTradeAddress",
        Text(Ram + "PostcodeCode", address.PostCode), Text(Ram + "LineOne", address.Line1), Text(Ram + "LineTwo", address.Line2), Text(Ram + "LineThree", address.Line3),
        Text(Ram + "CityName", address.City), Text(Ram + "CountryID", address.CountryCode), Text(Ram + "CountrySubDivisionName", address.Subdivision));

    private static IEnumerable<XElement> CiiPayments(IEnumerable<InvoicePayment> payments) {
        foreach (InvoicePayment payment in payments) {
            InvoiceBankAccount? account = payment.Account;
            yield return new XElement(Ram + "SpecifiedTradeSettlementPaymentMeans", Text(Ram + "TypeCode", payment.MeansCode), Text(Ram + "Information", payment.MeansText),
                payment.CardNumber == null ? null : new XElement(Ram + "ApplicableTradeSettlementFinancialCard", Text(Ram + "ID", payment.CardNumber), Text(Ram + "CardholderName", payment.CardHolder)),
                payment.DebitedAccount == null ? null : new XElement(Ram + "PayerPartyDebtorFinancialAccount", Text(Ram + "IBANID", payment.DebitedAccount)),
                account == null ? null : new XElement(Ram + "PayeePartyCreditorFinancialAccount", Text(Ram + (account.IsIban ? "IBANID" : "ProprietaryID"), account.Identifier), Text(Ram + "AccountName", account.Name)),
                account?.ProviderIdentifier == null ? null : new XElement(Ram + "PayeeSpecifiedCreditorFinancialInstitution", Text(Ram + "BICID", account.ProviderIdentifier)));
        }
    }
}
