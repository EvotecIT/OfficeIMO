namespace OfficeIMO.Invoicing;

public static partial class InvoiceModelValidator {
    private static bool HasTaxRegistration(InvoiceParty? party, string scheme) {
        InvoiceTaxRegistrationKind kind = scheme == InvoiceTaxRegistration.VatScheme
            ? InvoiceTaxRegistrationKind.Vat
            : InvoiceTaxRegistrationKind.Fiscal;
        return party?.TaxRegistrations.Any(registration => registration != null && registration.Kind == kind &&
            !string.IsNullOrWhiteSpace(registration.Identifier)) == true;
    }
}
