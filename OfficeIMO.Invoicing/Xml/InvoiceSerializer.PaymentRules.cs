namespace OfficeIMO.Invoicing;

public static partial class InvoiceSerializer {
    private static void CheckPaymentProfile(Invoice invoice, InvoiceXmlOptions options, Action<string, string> unsupported) {
        InvoicePayment payment = invoice.Payment!;
        if (options.Profile == InvoiceProfile.En16931) return;

        bool directDebit = payment.MeansCode == "49" || payment.MeansCode == "59";
        if (directDebit && string.IsNullOrWhiteSpace(payment.MandateReference))
            unsupported("Payment.MandateReference", "XRechnung and Peppol direct debit require a mandate reference for payment codes 49 and 59.");

        bool germanPaymentRules = options.Profile == InvoiceProfile.XRechnung ||
            invoice.Seller.Address?.CountryCode == "DE" && invoice.Buyer.Address?.CountryCode == "DE";
        if (!germanPaymentRules) return;

        // CII expresses BG-19 through separate fields; UBL uses PaymentMandate.
        bool hasDebitGroup = payment.MandateReference != null || payment.DebitedAccount != null ||
            options.Syntax == InvoiceSyntax.Cii && payment.CreditorIdentifier != null;
        if (hasDebitGroup || payment.MeansCode == "59") {
            if (string.IsNullOrWhiteSpace(payment.CreditorIdentifier))
                unsupported("Payment.CreditorIdentifier", "This profile's direct-debit group requires a bank-assigned creditor identifier.");
            if (string.IsNullOrWhiteSpace(payment.DebitedAccount))
                unsupported("Payment.DebitedAccount", "This profile's direct-debit group requires a debited account identifier.");
        }
        if (payment.MeansCode == "59") {
            if (payment.Accounts.Count != 0)
                unsupported("Payment.Accounts", "This profile's SEPA direct debit cannot include credit-transfer accounts.");
            if (payment.CardNumber != null)
                unsupported("Payment.CardNumber", "This profile's SEPA direct debit cannot include payment-card details.");
        }
    }
}
