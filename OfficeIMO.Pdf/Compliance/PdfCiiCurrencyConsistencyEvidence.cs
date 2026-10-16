namespace OfficeIMO.Pdf;

internal sealed class PdfCiiCurrencyConsistencyEvidence {
    internal PdfCiiCurrencyConsistencyEvidence(
        string? invoiceCurrencyCode, string? taxCurrencyCode,
        IReadOnlyList<string> amountCurrencyCodes,
        bool hasCurrencyAmount,
        IReadOnlyList<string> amountFieldsWithoutCurrency,
        IReadOnlyList<string> mismatchedAmountCurrencyFields) {
        InvoiceCurrencyCode = invoiceCurrencyCode;
        TaxCurrencyCode = taxCurrencyCode;
        AmountCurrencyCodes = amountCurrencyCodes;
        HasCurrencyAmount = hasCurrencyAmount;
        AmountFieldsWithoutCurrency = amountFieldsWithoutCurrency;
        MismatchedAmountCurrencyFields = mismatchedAmountCurrencyFields;
    }

    internal string? InvoiceCurrencyCode { get; }

    internal string? TaxCurrencyCode { get; }

    internal bool HasInvoiceCurrencyCode => !string.IsNullOrWhiteSpace(InvoiceCurrencyCode);

    internal IReadOnlyList<string> AmountCurrencyCodes { get; }

    internal bool HasCurrencyAmount { get; }

    internal IReadOnlyList<string> AmountFieldsWithoutCurrency { get; }

    internal IReadOnlyList<string> MismatchedAmountCurrencyFields { get; }

    internal bool AllAmountCurrenciesMatch => MismatchedAmountCurrencyFields.Count == 0;
}
