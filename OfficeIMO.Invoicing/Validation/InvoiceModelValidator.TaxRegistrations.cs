namespace OfficeIMO.Invoicing;

public static partial class InvoiceModelValidator {
    private static bool HasTaxRegistration(InvoiceParty? party, string scheme) => party?.TaxRegistrations.Any(registration =>
        registration != null && string.Equals(registration.SchemeId, scheme, StringComparison.Ordinal) &&
        !string.IsNullOrWhiteSpace(registration.Identifier)) == true;
}
