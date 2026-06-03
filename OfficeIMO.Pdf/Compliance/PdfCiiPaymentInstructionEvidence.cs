namespace OfficeIMO.Pdf;

internal sealed class PdfCiiPaymentInstructionEvidence {
    internal PdfCiiPaymentInstructionEvidence(
        bool hasSpecifiedTradeSettlementPaymentMeans,
        bool hasTypeCode,
        bool hasCreditorFinancialAccount,
        bool hasCreditorAccountId) {
        HasSpecifiedTradeSettlementPaymentMeans = hasSpecifiedTradeSettlementPaymentMeans;
        HasTypeCode = hasTypeCode;
        HasCreditorFinancialAccount = hasCreditorFinancialAccount;
        HasCreditorAccountId = hasCreditorAccountId;
    }

    internal bool HasSpecifiedTradeSettlementPaymentMeans { get; }

    internal bool HasTypeCode { get; }

    internal bool HasCreditorFinancialAccount { get; }

    internal bool HasCreditorAccountId { get; }
}
