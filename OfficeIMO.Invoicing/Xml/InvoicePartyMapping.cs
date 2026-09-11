namespace OfficeIMO.Invoicing;

/// <summary>One role contract for authored parties and explicitly loss-approved parsed parties.</summary>
internal static class InvoicePartyMapping {
    internal static void Check(InvoiceParty? party, string role, Action<string, string> unsupported, bool discardUnsupported = false) {
        if (party == null || role == "Seller") return;
        void Reject(string field, Action discard) {
            unsupported(role + "." + field, "The supported " + role + " mapping does not carry " + field + ".");
            if (discardUnsupported) discard();
        }
        if (party.LegalInformation != null) Reject("LegalInformation", () => party.LegalInformation = null);
        if (party.TaxRegistration != null) Reject("TaxRegistration", () => party.TaxRegistration = null);
        if (role == "Buyer") return;
        if (party.TradingName != null) Reject("TradingName", () => party.TradingName = null);
        if (party.Contact != null) Reject("Contact", () => party.Contact = null);
        if (party.ElectronicAddress != null) Reject("ElectronicAddress", () => party.ElectronicAddress = null);
        if (role == "Payee") {
            if (party.VatIdentifier != null) Reject("VatIdentifier", () => party.VatIdentifier = null);
            InvoiceAddress? address = party.Address;
            if (address != null && (!string.IsNullOrEmpty(address.CountryCode) || address.Line1 != null || address.Line2 != null || address.Line3 != null || address.City != null || address.PostCode != null || address.Subdivision != null))
                Reject("Address", () => party.Address = new InvoiceAddress());
            if (party.Identifiers.Count > 1)
                Reject("Identifiers", () => { while (party.Identifiers.Count > 1) party.Identifiers.RemoveAt(party.Identifiers.Count - 1); });
        } else if (role == "TaxRepresentative") {
            if (party.Identifiers.Count != 0) Reject("Identifiers", () => party.Identifiers.Clear());
            if (party.LegalRegistration != null) Reject("LegalRegistration", () => party.LegalRegistration = null);
        }
    }
}
