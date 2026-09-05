namespace OfficeIMO.Invoices.Internal;

internal sealed class CiiCurrencyConsistencyEvidence {
    internal CiiCurrencyConsistencyEvidence(
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

    internal bool AllAmountCurrenciesMatch => MismatchedAmountCurrencyFields.Count == 0;
}
