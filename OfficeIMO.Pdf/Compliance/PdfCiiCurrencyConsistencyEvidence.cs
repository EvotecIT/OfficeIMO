namespace OfficeIMO.Pdf;

internal sealed class PdfCiiCurrencyConsistencyEvidence {
    internal PdfCiiCurrencyConsistencyEvidence(
        string? invoiceCurrencyCode,
        IReadOnlyList<string> amountCurrencyCodes,
        bool hasCurrencyAmount,
        IReadOnlyList<string> amountFieldsWithoutCurrency,
        IReadOnlyList<string> mismatchedAmountCurrencyFields) {
        InvoiceCurrencyCode = invoiceCurrencyCode;
        AmountCurrencyCodes = amountCurrencyCodes;
        HasCurrencyAmount = hasCurrencyAmount;
        AmountFieldsWithoutCurrency = amountFieldsWithoutCurrency;
        MismatchedAmountCurrencyFields = mismatchedAmountCurrencyFields;
    }

    internal string? InvoiceCurrencyCode { get; }

    internal bool HasInvoiceCurrencyCode => !string.IsNullOrWhiteSpace(InvoiceCurrencyCode);

    internal IReadOnlyList<string> AmountCurrencyCodes { get; }

    internal bool HasCurrencyAmount { get; }

    internal IReadOnlyList<string> AmountFieldsWithoutCurrency { get; }

    internal IReadOnlyList<string> MismatchedAmountCurrencyFields { get; }

    internal bool AllAmountCurrenciesMatch => AmountFieldsWithoutCurrency.Count == 0 && MismatchedAmountCurrencyFields.Count == 0;
}
