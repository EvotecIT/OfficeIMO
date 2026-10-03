using System.Globalization;

namespace OfficeIMO.Invoicing;

public static partial class Fa3InvoiceWriter {
    private static void CheckUnsupportedInvoiceFields(Invoice invoice, Action<string, object?> optional, Action<string, string> error) {
        optional("BusinessProcessId", invoice.BusinessProcessId); optional("BuyerReference", invoice.BuyerReference);
        optional("Payee", invoice.Payee); optional("TaxRepresentative", invoice.TaxRepresentative); optional("Delivery", invoice.Delivery);
        optional("TaxPointDateCode", invoice.TaxPointDateCode); optional("ProjectReference", invoice.ProjectReference);
        optional("ContractReference", invoice.ContractReference); optional("PurchaseOrderReference", invoice.PurchaseOrderReference);
        optional("SalesOrderReference", invoice.SalesOrderReference); optional("ReceivingAdviceReference", invoice.ReceivingAdviceReference);
        optional("DespatchAdviceReference", invoice.DespatchAdviceReference); optional("TenderReference", invoice.TenderReference);
        optional("AccountingReference", invoice.AccountingReference); optional("ObjectIdentifier", invoice.ObjectIdentifier);
        optional("PaymentReference", invoice.PaymentReference); optional("CreditorIdentifier", invoice.CreditorIdentifier);
        optional("DirectDebitMandateReference", invoice.DirectDebitMandateReference); optional("PaymentTerms", invoice.PaymentTerms);
        optional("DeclaredTotals", invoice.DeclaredTotals);
        if (invoice.DeclaredTaxes.Count != 0) error("DeclaredTaxes", "Use explicit national FiscalAmounts; EN declarations cannot be relabeled as national buckets.");
        if (invoice.Notes.Count != 0) error("Notes", "Document notes have no supported national authoring mapping.");
        if (invoice.SupportingDocuments.Count != 0) error("SupportingDocuments", "Supporting documents require a national attachment contract.");
        if (invoice.AllowancesAndCharges.Count != 0) error("AllowancesAndCharges", "Document adjustments require an explicit national settlement contract.");
        if (invoice.PrepaidAmount != 0 || invoice.RoundingAmount != 0) error("PrepaidAmount/RoundingAmount", "EN prepayment and payable rounding cannot be relabeled as Polish payment or settlement records.");
        if (invoice.Payments.Count > 100) error("Payments", "At most 100 accounts are supported.");
        string? means = null;
        for (int index = 0; index < invoice.Payments.Count; index++) {
            InvoicePayment? payment = invoice.Payments[index]; string path = "Payments[" + index + "]";
            if (payment == null) { error(path, "Payment is null."); continue; }
            if (payment.MeansCode == "58" || PaymentForm(payment.MeansCode) == null) error(path + ".MeansCode", "Choose an explicit supported national payment mapping: common cash 10, card 48, cheque 20 or generic transfer 30. SEPA-specific classification cannot be preserved by FormaPlatnosci.");
            if (means != null && means != payment.MeansCode) error(path + ".MeansCode", "FA(3) permits one national payment form for these instructions.");
            means = payment.MeansCode;
            optional(path + ".MeansText", payment.MeansText); optional(path + ".Reference", payment.Reference);
            optional(path + ".CardNumber", payment.CardNumber); optional(path + ".CardNetworkId", payment.CardNetworkId);
            optional(path + ".CardHolder", payment.CardHolder); optional(path + ".MandateReference", payment.MandateReference);
            optional(path + ".CreditorIdentifier", payment.CreditorIdentifier); optional(path + ".DebitedAccount", payment.DebitedAccount);
            if (payment.Account == null) continue;
            if (payment.MeansCode != "30") error(path + ".Account", "A creditor account requires the supported generic-transfer mapping.");
            optional(path + ".Account.Name", payment.Account.Name);
            string account = payment.Account.Identifier;
            if (string.IsNullOrWhiteSpace(account) || account.Length > 34 || account.Any(character => character < '0' || character > '9') && !InvoiceBankAccountIdentity.IsValidIban(account)) error(path + ".Account.Identifier", "Supply a valid IBAN or a numeric national account identifier of at most 34 characters.");
            if (payment.Account.IsIban && !InvoiceBankAccountIdentity.IsValidIban(account)) error(path + ".Account.IsIban", "An account marked as an IBAN must have a valid registered structure and checksum.");
            string? bic = payment.Account.ProviderIdentifier;
            if (bic != null && (bic.Length is not (8 or 11) || bic.Any(character => !(character >= 'A' && character <= 'Z' || character >= '0' && character <= '9')))) error(path + ".Account.ProviderIdentifier", "SWIFT must contain eight or eleven uppercase alphanumeric characters.");
        }
    }

    private static void CheckParty(InvoiceParty party, string path, bool seller, Fa3InvoiceKind kind,
        Action<string, string?, int, bool> text, Action<string, string> error, Action<string, object?> optional) {
        if (seller || kind != Fa3InvoiceKind.Simplified || !string.IsNullOrEmpty(party.Name)) text(path + ".Name", party.Name, 512, true);
        optional(path + ".TradingName", party.TradingName); optional(path + ".LegalInformation", party.LegalInformation);
        optional(path + ".LegalRegistration", party.LegalRegistration); optional(path + ".ElectronicAddress", party.ElectronicAddress);
        if (party.Identifiers.Count != 0) error(path + ".Identifiers", "Business identifiers are distinct from the supported NIP tax identity.");
        if (party.TaxRegistrations.Count > 1 || seller && party.TaxRegistrations.Count != 1) error(path + ".TaxRegistrations", "Supply one seller NIP and at most one buyer NIP; additional identities cannot be discarded.");
        foreach (InvoiceTaxRegistration registration in party.TaxRegistrations) {
            if (registration == null || !(registration.Kind == InvoiceTaxRegistrationKind.Fiscal && registration.SchemeId == "NIP" ||
                registration.Kind == InvoiceTaxRegistrationKind.Vat && registration.SchemeId == InvoiceTaxRegistration.VatScheme && registration.Identifier?.StartsWith("PL", StringComparison.Ordinal) == true)) {
                error(path + ".TaxRegistrations", "Supply an explicit NIP fiscal registration or Polish VAT registration; foreign identification requires a separate national mapping."); continue;
            }
            string? nip = registration.Kind == InvoiceTaxRegistrationKind.Vat ? registration.Identifier?.Substring(2) : registration.Identifier;
            if (nip == null || nip.Length != 10 || !System.Text.RegularExpressions.Regex.IsMatch(nip, "^[1-9]([0-9][1-9]|[1-9][0-9])[0-9]{7}$")) error(path + ".TaxRegistrations", "NIP must match the pinned ten-digit schema pattern; this does not verify tax registration.");
        }
        InvoiceAddress address = party.Address;
        text(path + ".Address.Line1", address.Line1, 512, seller);
        text(path + ".Address.Line2", address.Line2, 512, false);
        if (address.Line1 != null) {
            text(path + ".Address.CountryCode", address.CountryCode, 2, true);
            if (address.CountryCode == null || !Fa3ScalarContract.Countries.Contains(address.CountryCode)) error(path + ".Address.CountryCode", "Supply a country code in the pinned national schema dictionary.");
        } else if (address.Line2 != null || !string.IsNullOrEmpty(address.CountryCode)) error(path + ".Address", "Country and second address line require a first address line.");
        optional(path + ".Address.Line3", address.Line3); optional(path + ".Address.City", address.City);
        optional(path + ".Address.PostCode", address.PostCode); optional(path + ".Address.Subdivision", address.Subdivision);
        if (party.Contact != null) {
            optional(path + ".Contact.Name", party.Contact.Name);
            text(path + ".Contact.Email", party.Contact.Email, 255, false); text(path + ".Contact.Telephone", party.Contact.Telephone, 16, false);
            if (party.Contact.Email is { Length: <= 255 } email &&
                !System.Text.RegularExpressions.Regex.IsMatch(email, "^.+@.+$"))
                error(path + ".Contact.Email", "Email must match the pinned schema pattern.");
        }
    }

    private static void CheckLines(IEnumerable<InvoiceLine> lines, string path, Fa3InvoiceWriteOptions options,
        Action<string, string?, int, bool> text, Action<string, string> error, Action<string, object?> optional, Action<string, decimal> money) {
        int count = 0; var ids = new HashSet<string>(StringComparer.Ordinal);
        foreach (InvoiceLine line in lines) {
            string location = path + "[" + count++ + "]";
            if (count > 10_000) { error(path, "At most 10,000 rows are supported."); break; }
            if (line == null || line.Tax == null) { error(location, "Line and tax must not be null."); continue; }
            if (string.IsNullOrEmpty(line.Id) || string.IsNullOrEmpty(line.UnitCode)) { error(location, "Line identifier and unit code are required."); continue; }
            if (!int.TryParse(line.Id, NumberStyles.None, CultureInfo.InvariantCulture, out int id) || id <= 0 || !ids.Add(line.Id)) error(location + ".Id", "Supply a distinct positive integer row identifier.");
            text(location + ".Name", line.Name, 512, true); text(location + ".UnitCode", line.UnitCode, 16, true);
            string? unit = UnitLabel(line.UnitCode, options);
            text(location + ".UnitLabel", unit, 256, true);
            if (decimal.Round(line.Quantity, 6) != line.Quantity || line.Quantity > 9_999_999_999_999_999.999999m || line.Quantity < -9_999_999_999_999_999.999999m) error(location + ".Quantity", "Quantity exceeds the national six-fractional-digit contract.");
            if (line.UnitPrice < 0 || decimal.Round(line.UnitPrice, 8) != line.UnitPrice || line.UnitPrice > 99_999_999_999_999.99999999m) error(location + ".UnitPrice", "Price requires a nonnegative national value with at most eight fractional digits.");
            if (line.PriceBaseQuantity != 1) error(location + ".PriceBaseQuantity", "The supported national price applies to one unit; do not discard a price base quantity.");
            optional(location + ".GrossPrice", line.GrossPrice); optional(location + ".PriceDiscount", line.PriceDiscount);
            optional(location + ".Description", line.Description); optional(location + ".Note", line.Note); optional(location + ".Period", line.Period);
            optional(location + ".OrderLineReference", line.OrderLineReference); optional(location + ".AccountingReference", line.AccountingReference);
            optional(location + ".ObjectIdentifier", line.ObjectIdentifier); optional(location + ".BuyerItemIdentifier", line.BuyerItemIdentifier);
            optional(location + ".StandardItemIdentifier", line.StandardItemIdentifier); optional(location + ".OriginCountryCode", line.OriginCountryCode);
            text(location + ".SellerItemIdentifier", line.SellerItemIdentifier, 50, false);
            if (line.AllowancesAndCharges.Count != 0 || line.Classifications.Count != 0 || line.Attributes.Count != 0) error(location, "Line adjustments, classifications and attributes require additional national mappings.");
            if (line.DeclaredNetAmount.HasValue) money(location + ".DeclaredNetAmount", line.DeclaredNetAmount.Value);
            optional(location + ".Tax.ExemptionReasonCode", line.Tax.ExemptionReasonCode);
            if (line.Tax.ExemptionReason != null && line.Tax.ExemptionReason != options.Annotations.ExemptionLegalBasis) error(location + ".Tax.ExemptionReason", "Exemption text must agree with the explicitly supplied national legal basis.");
            if (options.Annotations.MarginProcedure != Fa3MarginProcedure.None) {
                if (line.Tax.Code != "O" || line.Tax.Rate.HasValue) error(location + ".Tax", "Margin rows require the common outside-scope category without a VAT rate and explicit national fiscal declarations.");
                continue;
            }
            string? label = TaxLabel(line, options), suffix = TaxSuffix(label);
            if (suffix == null) error(location + ".Tax", "Supply a supported national rate label. Outside-territory categories require explicit np I or np II.");
            string expectedCategory = suffix switch { "1" or "2" or "3" or "4" => "S", "6_1" => "Z", "6_2" => "K", "6_3" => "G", "7" => "E", "8" or "9" => "O", "10" => "AE", _ => string.Empty };
            if (line.Tax.Code != expectedCategory || expectedCategory == "S" && label != (line.Tax.Rate.HasValue ? Number(line.Tax.Rate.Value) : null) || expectedCategory != "S" && expectedCategory != "O" && line.Tax.Rate != 0 || expectedCategory == "O" && line.Tax.Rate.HasValue) error(location + ".Tax", "The explicit national label must agree with the common category and rate.");
            if (expectedCategory == "E" && options.Annotations.ExemptionLegalBasis == null) error("Annotations.ExemptionLegalBasis", "Exempt rows require an explicit national legal basis.");
            if (expectedCategory == "AE" && !options.Annotations.ReverseCharge) error("Annotations.ReverseCharge", "Reverse-charge rows require the national reverse-charge declaration.");
        }
    }
}
