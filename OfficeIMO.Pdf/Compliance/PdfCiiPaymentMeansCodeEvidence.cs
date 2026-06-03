namespace OfficeIMO.Pdf;

internal sealed class PdfCiiPaymentMeansCodeEvidence {
    internal PdfCiiPaymentMeansCodeEvidence(
        bool hasSpecifiedTradeSettlementPaymentMeans,
        bool hasTypeCode,
        IReadOnlyList<string> typeCodes) {
        HasSpecifiedTradeSettlementPaymentMeans = hasSpecifiedTradeSettlementPaymentMeans;
        HasTypeCode = hasTypeCode;
        TypeCodes = typeCodes;
    }

    internal bool HasSpecifiedTradeSettlementPaymentMeans { get; }

    internal bool HasTypeCode { get; }

    internal IReadOnlyList<string> TypeCodes { get; }
}
