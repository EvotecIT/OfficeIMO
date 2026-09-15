namespace OfficeIMO.Invoicing;

public static partial class InvoiceSerializer {
    private static void CheckPaymentProfile(Invoice invoice, InvoiceXmlOptions options, Action<string, string> unsupported) {
        for (int index = 0; index < invoice.Payments.Count; index++) {
            InvoicePayment payment = invoice.Payments[index];
            string path = "Payments[" + index + "]";

            if (options.Syntax == InvoiceSyntax.Cii && payment.CardNetworkId != null)
                unsupported(path + ".CardNetworkId", "CII has no target field for card network identifier '" + payment.CardNetworkId + "'.");

            if (options.Profile == InvoiceProfile.En16931 || options.Release == InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2) continue;

            bool directDebit = payment.MeansCode == "49" || payment.MeansCode == "59";
            if (directDebit && string.IsNullOrWhiteSpace(payment.MandateReference))
                unsupported(path + ".MandateReference", "XRechnung and Peppol direct debit require a mandate reference for payment codes 49 and 59.");

            bool germanPaymentRules = options.Profile == InvoiceProfile.XRechnung ||
                invoice.Seller.Address?.CountryCode == "DE" && invoice.Buyer.Address?.CountryCode == "DE";
            if (!germanPaymentRules) continue;

            bool hasDebitGroup = payment.MandateReference != null || payment.DebitedAccount != null ||
                options.Syntax == InvoiceSyntax.Cii && payment.CreditorIdentifier != null;
            if (hasDebitGroup || payment.MeansCode == "59") {
                if (string.IsNullOrWhiteSpace(payment.CreditorIdentifier))
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
