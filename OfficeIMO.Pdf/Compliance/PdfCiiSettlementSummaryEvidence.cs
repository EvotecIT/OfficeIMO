namespace OfficeIMO.Pdf;

internal sealed class PdfCiiSettlementSummaryEvidence {
    internal PdfCiiSettlementSummaryEvidence(
        bool hasApplicableHeaderTradeSettlement,
        bool hasInvoiceCurrencyCode,
        bool hasApplicableTradeTax,
        bool hasTaxBasisTotalAmount,
        bool hasTaxTotalAmount) {
        HasApplicableHeaderTradeSettlement = hasApplicableHeaderTradeSettlement;
        HasInvoiceCurrencyCode = hasInvoiceCurrencyCode;
        HasApplicableTradeTax = hasApplicableTradeTax;
        HasTaxBasisTotalAmount = hasTaxBasisTotalAmount;
        HasTaxTotalAmount = hasTaxTotalAmount;
    }

    internal bool HasApplicableHeaderTradeSettlement { get; }

    internal bool HasInvoiceCurrencyCode { get; }

    internal bool HasApplicableTradeTax { get; }

    internal bool HasTaxBasisTotalAmount { get; }

    internal bool HasTaxTotalAmount { get; }
}
