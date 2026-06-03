namespace OfficeIMO.Pdf;

internal sealed class PdfCiiPaymentInstructionEvidence {
    internal PdfCiiPaymentInstructionEvidence(
        bool hasSpecifiedTradeSettlementPaymentMeans,
        bool hasTypeCode,
        bool hasCreditorFinancialAccount,
        bool hasCreditorAccountId,
        IReadOnlyList<string> typeCodes,
        IReadOnlyList<string> missingTypeCodePaymentMeans,
        IReadOnlyList<string> missingCreditorAccountPaymentMeans,
        IReadOnlyList<string> missingCreditorAccountIdPaymentMeans) {
        HasSpecifiedTradeSettlementPaymentMeans = hasSpecifiedTradeSettlementPaymentMeans;
        HasTypeCode = hasTypeCode;
        HasCreditorFinancialAccount = hasCreditorFinancialAccount;
        HasCreditorAccountId = hasCreditorAccountId;
        TypeCodes = typeCodes;
        MissingTypeCodePaymentMeans = missingTypeCodePaymentMeans;
        MissingCreditorAccountPaymentMeans = missingCreditorAccountPaymentMeans;
        MissingCreditorAccountIdPaymentMeans = missingCreditorAccountIdPaymentMeans;
    }

    internal bool HasSpecifiedTradeSettlementPaymentMeans { get; }

    internal bool HasTypeCode { get; }

    internal bool HasCreditorFinancialAccount { get; }

    internal bool HasCreditorAccountId { get; }

    internal IReadOnlyList<string> TypeCodes { get; }

    internal IReadOnlyList<string> MissingTypeCodePaymentMeans { get; }

    internal IReadOnlyList<string> MissingCreditorAccountPaymentMeans { get; }

    internal IReadOnlyList<string> MissingCreditorAccountIdPaymentMeans { get; }
}
