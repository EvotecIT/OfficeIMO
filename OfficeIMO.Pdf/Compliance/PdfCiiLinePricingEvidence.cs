namespace OfficeIMO.Pdf;

internal sealed class PdfCiiLinePricingEvidence {
    internal PdfCiiLinePricingEvidence(
        bool hasIncludedSupplyChainTradeLineItem,
        bool hasSpecifiedLineTradeAgreement,
        bool hasProductTradePrice,
        bool hasPriceChargeAmount) {
        HasIncludedSupplyChainTradeLineItem = hasIncludedSupplyChainTradeLineItem;
        HasSpecifiedLineTradeAgreement = hasSpecifiedLineTradeAgreement;
        HasProductTradePrice = hasProductTradePrice;
        HasPriceChargeAmount = hasPriceChargeAmount;
    }

    internal bool HasIncludedSupplyChainTradeLineItem { get; }

    internal bool HasSpecifiedLineTradeAgreement { get; }

    internal bool HasProductTradePrice { get; }

    internal bool HasPriceChargeAmount { get; }
}
