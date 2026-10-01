namespace OfficeIMO.Invoicing;

public static partial class InvoiceSerializer {
    private static void CheckIndependentPaymentData(Invoice invoice, InvoiceXmlOptions options, Action<string, string> unsupported) {
        void Conflict(string? header, Func<InvoicePayment, string?> selector, string field) {
            if (header == null) return;
            for (int index = 0; index < invoice.Payments.Count; index++) {
                string? value = selector(invoice.Payments[index]);
                if (value != null && value != header)
                    unsupported("Payments[" + index + "]." + field, "Payment value '" + value + "' conflicts with invoice-level value '" + header + "'.");
            }
        }
        Conflict(invoice.CreditorIdentifier, payment => payment.CreditorIdentifier, "CreditorIdentifier");
        Conflict(invoice.PaymentReference, payment => payment.Reference, "Reference");
        Conflict(invoice.DirectDebitMandateReference, payment => payment.MandateReference, "MandateReference");
        if (options.Syntax == InvoiceSyntax.Ubl && invoice.Payments.Count == 0 && invoice.PaymentReference != null)
            unsupported("PaymentReference", "UBL requires the remittance reference to belong to an explicit payment instruction. Supply its means code without inventing one during conversion.");
        if (options.Syntax == InvoiceSyntax.Ubl && invoice.Payments.Count == 0 && invoice.DirectDebitMandateReference != null)
            unsupported("DirectDebitMandateReference", "UBL requires the mandate to belong to an explicit payment instruction.");
        if (options.Syntax == InvoiceSyntax.Cii && invoice.Payments.Count == 0 && RequiresGermanPaymentRules(invoice, options) &&
            (invoice.DirectDebitMandateReference != null || invoice.CreditorIdentifier != null)) {
            if (string.IsNullOrWhiteSpace(invoice.CreditorIdentifier))
                unsupported("CreditorIdentifier", "This profile's direct-debit group requires a bank-assigned creditor identifier.");
            unsupported("Payments", "This profile's invoice-level direct-debit data requires a debited account on an explicit payment instruction.");
        }
    }
    private static bool RequiresGermanPaymentRules(Invoice invoice, InvoiceXmlOptions options) =>
        options.Release != InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2 &&
        (options.Profile == InvoiceProfile.XRechnung || options.Profile == InvoiceProfile.PeppolBis &&
            invoice.Seller.Address?.CountryCode == "DE" && invoice.Buyer.Address?.CountryCode == "DE");
    private static void CheckPaymentProfile(Invoice invoice, InvoiceXmlOptions options, Action<string, string> unsupported) {
        string? ciiMandate = invoice.DirectDebitMandateReference ?? FirstPaymentValue(invoice, value => value.MandateReference);
        string? creditor = invoice.CreditorIdentifier ?? FirstPaymentValue(invoice, value => value.CreditorIdentifier);
        for (int index = 0; index < invoice.Payments.Count; index++) {
            InvoicePayment payment = invoice.Payments[index];
            string path = "Payments[" + index + "]";

            if (options.Syntax == InvoiceSyntax.Cii && payment.CardNetworkId != null)
                unsupported(path + ".CardNetworkId", "CII has no target field for card network identifier '" + payment.CardNetworkId + "'.");

            if (options.Profile == InvoiceProfile.En16931 || options.Release == InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2) continue;

            bool directDebit = payment.MeansCode == "49" || payment.MeansCode == "59";
            string? mandate = options.Syntax == InvoiceSyntax.Cii
                ? ciiMandate
                : payment.MandateReference ?? invoice.DirectDebitMandateReference;
            if (directDebit && string.IsNullOrWhiteSpace(mandate))
                unsupported(path + ".MandateReference", "XRechnung and Peppol direct debit require a mandate reference for payment codes 49 and 59.");

            if (!RequiresGermanPaymentRules(invoice, options)) continue;

            bool hasDebitGroup = mandate != null || payment.DebitedAccount != null ||
                options.Syntax == InvoiceSyntax.Cii && creditor != null;
            if (hasDebitGroup || payment.MeansCode == "59") {
                if (string.IsNullOrWhiteSpace(creditor))
                    unsupported(path + ".CreditorIdentifier", "This profile's direct-debit group requires a bank-assigned creditor identifier.");
                if (string.IsNullOrWhiteSpace(payment.DebitedAccount))
                    unsupported(path + ".DebitedAccount", "This profile's direct-debit group requires a debited account identifier.");
            }
            if (payment.MeansCode == "59") {
                if (payment.Account != null)
                    unsupported(path + ".Account", "This profile's SEPA direct debit cannot include a credit-transfer account.");
                if (payment.CardNumber != null)
                    unsupported(path + ".CardNumber", "This profile's SEPA direct debit cannot include payment-card details.");
            }
        }
    }
}
