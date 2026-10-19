namespace OfficeIMO.Pdf;

internal static partial class PdfCiiDocumentHeaderInspector {
    // Accounting-currency tax must never replace the invoice-currency amount in
    // payable or VAT-breakdown arithmetic, regardless of XML occurrence order.
    private static bool TryReadInvoiceCurrencyTaxTotal(System.Xml.XmlReader reader, string? invoiceCurrency,
        HashSet<string> seenCurrencies, ref string? diagnostic, out decimal? amount) {
        string currency = reader.GetAttribute("currencyID") ?? string.Empty;
        if (!seenCurrencies.Add(currency)) diagnostic = "Duplicate TaxTotalAmount for currency " + currency + ".";
        bool parsed = TryReadAmount(reader, "TaxTotalAmount", ref diagnostic, out amount);
        return parsed && (invoiceCurrency == null || string.Equals(currency, invoiceCurrency, StringComparison.Ordinal));
    }
}
