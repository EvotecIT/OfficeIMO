namespace OfficeIMO.Invoicing;

public static partial class InvoiceModelValidator {
    private sealed partial class ModelChecks {
        internal void TaxRequirements(Invoice invoice, InvoiceCalculation calculation) {
            // Check the emitted breakdown: imported reasons may exist only on DeclaredTaxes.
            bool sellerVat = !string.IsNullOrWhiteSpace(invoice.Seller?.VatIdentifier);
            bool representativeVat = !string.IsNullOrWhiteSpace(invoice.TaxRepresentative?.VatIdentifier);
            bool sellerTax = sellerVat || representativeVat || !string.IsNullOrWhiteSpace(invoice.Seller?.TaxRegistration);
            bool buyerVat = !string.IsNullOrWhiteSpace(invoice.Buyer?.VatIdentifier);
            foreach (InvoiceCalculatedTax tax in calculation.Taxes) {
                string path = "Taxes[" + tax.CategoryCode + "]";
                bool exempt = new[] { "E", "AE", "G", "K", "O" }.Contains(tax.CategoryCode);
                bool hasReason = !string.IsNullOrWhiteSpace(tax.ExemptionReason) || !string.IsNullOrWhiteSpace(tax.ExemptionReasonCode);
                if (exempt && !hasReason)
                    Error("INV-VAT-EXEMPTION", "This VAT category requires an exemption reason text or code.", path);
                if (!exempt && (tax.ExemptionReason != null || tax.ExemptionReasonCode != null))
                    Error("INV-VAT-EXEMPTION", "This VAT category must not carry an exemption reason.", path);
                if (new[] { "Z", "E", "AE", "G", "K", "O" }.Contains(tax.CategoryCode) && tax.TaxAmount != 0m)
                    Error("INV-VAT-ZERO", "This VAT category requires an exactly zero tax amount.", path);
                if (tax.CategoryCode == "O") {
                    if (sellerVat || representativeVat || buyerVat)
                        Error("INV-VAT-IDENTIFIER", "Outside-scope VAT must not carry seller, representative or buyer VAT identifiers.", path);
                    if (calculation.Taxes.Count != 1)
                        Error("INV-VAT-MIX", "Outside-scope VAT cannot be combined with other VAT categories.", path);
                } else if (tax.CategoryCode == "G" || tax.CategoryCode == "K") {
                    if (!sellerVat && !representativeVat)
                        Error("INV-VAT-IDENTIFIER", "This VAT category requires a seller or tax representative VAT identifier.", "Seller.VatIdentifier");
                } else if (!sellerTax) {
                    Error("INV-VAT-IDENTIFIER", "Supply a seller VAT/tax registration or tax representative VAT identifier.", "Seller");
                }
                if (tax.CategoryCode == "AE" && !buyerVat && string.IsNullOrWhiteSpace(invoice.Buyer?.LegalRegistration?.Value))
                    Error("INV-VAT-IDENTIFIER", "Reverse charge requires a buyer VAT or legal registration identifier.", "Buyer");
                if (tax.CategoryCode == "K") {
                    if (!buyerVat) Error("INV-VAT-IDENTIFIER", "Intra-community supply requires a buyer VAT identifier.", "Buyer.VatIdentifier");
                    if (invoice.Delivery?.Date == null && invoice.Period?.Start == null && invoice.Period?.End == null)
                        Error("INV-VAT-DELIVERY", "Intra-community supply requires a delivery date or invoicing period.", "Delivery.Date");
                    if (string.IsNullOrWhiteSpace(invoice.Delivery?.Address?.CountryCode))
                        Error("INV-VAT-DELIVERY", "Intra-community supply requires a deliver-to country.", "Delivery.Address.CountryCode");
                }
            }
        }
    }
}
