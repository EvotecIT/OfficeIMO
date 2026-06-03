namespace OfficeIMO.Pdf;

internal sealed class PdfCiiPaymentAccountEvidence {
    internal PdfCiiPaymentAccountEvidence(
        bool hasSpecifiedTradeSettlementPaymentMeans,
        bool hasCreditorFinancialAccount,
        bool hasAccountId,
        bool hasIbanId,
        bool allIbanIdsAreValid,
        IReadOnlyList<string> invalidIbanIds) {
        HasSpecifiedTradeSettlementPaymentMeans = hasSpecifiedTradeSettlementPaymentMeans;
        HasCreditorFinancialAccount = hasCreditorFinancialAccount;
        HasAccountId = hasAccountId;
        HasIbanId = hasIbanId;
        AllIbanIdsAreValid = allIbanIdsAreValid;
        InvalidIbanIds = invalidIbanIds;
    }

    internal bool HasSpecifiedTradeSettlementPaymentMeans { get; }

    internal bool HasCreditorFinancialAccount { get; }

    internal bool HasAccountId { get; }

    internal bool HasIbanId { get; }

    internal bool AllIbanIdsAreValid { get; }

    internal IReadOnlyList<string> InvalidIbanIds { get; }
}
