namespace OfficeIMO.Invoices.Internal;

internal sealed class CiiSettlementSummaryEvidence {
    internal CiiSettlementSummaryEvidence(
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
