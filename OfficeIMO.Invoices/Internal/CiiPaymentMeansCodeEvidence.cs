namespace OfficeIMO.Invoices.Internal;

internal sealed class CiiPaymentMeansCodeEvidence {
    internal CiiPaymentMeansCodeEvidence(
        bool hasSpecifiedTradeSettlementPaymentMeans,
        bool hasTypeCode,
        IReadOnlyList<string> typeCodes,
        IReadOnlyList<string> missingTypeCodePaymentMeans) {
        HasSpecifiedTradeSettlementPaymentMeans = hasSpecifiedTradeSettlementPaymentMeans;
        HasTypeCode = hasTypeCode;
        TypeCodes = typeCodes;
        MissingTypeCodePaymentMeans = missingTypeCodePaymentMeans;
    }

    internal bool HasSpecifiedTradeSettlementPaymentMeans { get; }

    internal bool HasTypeCode { get; }

    internal IReadOnlyList<string> TypeCodes { get; }

    internal IReadOnlyList<string> MissingTypeCodePaymentMeans { get; }
}
