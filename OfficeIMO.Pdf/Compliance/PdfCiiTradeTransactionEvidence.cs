namespace OfficeIMO.Pdf;

internal sealed class PdfCiiTradeTransactionEvidence {
    internal PdfCiiTradeTransactionEvidence(
        bool hasSupplyChainTradeTransaction,
        bool hasApplicableHeaderTradeAgreement,
        bool hasSellerTradeParty,
        bool hasBuyerTradeParty,
        bool hasApplicableHeaderTradeSettlement,
        bool hasSpecifiedTradeSettlementHeaderMonetarySummation,
        bool hasPayableOrGrandTotalAmount) {
        HasSupplyChainTradeTransaction = hasSupplyChainTradeTransaction;
        HasApplicableHeaderTradeAgreement = hasApplicableHeaderTradeAgreement;
        HasSellerTradeParty = hasSellerTradeParty;
        HasBuyerTradeParty = hasBuyerTradeParty;
        HasApplicableHeaderTradeSettlement = hasApplicableHeaderTradeSettlement;
        HasSpecifiedTradeSettlementHeaderMonetarySummation = hasSpecifiedTradeSettlementHeaderMonetarySummation;
        HasPayableOrGrandTotalAmount = hasPayableOrGrandTotalAmount;
    }

    internal bool HasSupplyChainTradeTransaction { get; }

    internal bool HasApplicableHeaderTradeAgreement { get; }

    internal bool HasSellerTradeParty { get; }

    internal bool HasBuyerTradeParty { get; }

    internal bool HasApplicableHeaderTradeSettlement { get; }

    internal bool HasSpecifiedTradeSettlementHeaderMonetarySummation { get; }

    internal bool HasPayableOrGrandTotalAmount { get; }
}
