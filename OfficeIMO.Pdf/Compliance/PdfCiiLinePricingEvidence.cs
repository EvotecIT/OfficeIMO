namespace OfficeIMO.Pdf;

internal sealed class PdfCiiLinePricingEvidence {
    internal PdfCiiLinePricingEvidence(
        bool hasIncludedSupplyChainTradeLineItem,
        bool hasSpecifiedLineTradeAgreement,
        bool hasProductTradePrice,
        bool hasPriceChargeAmount,
        IReadOnlyList<string> missingLinePricingFields) {
        HasIncludedSupplyChainTradeLineItem = hasIncludedSupplyChainTradeLineItem;
        HasSpecifiedLineTradeAgreement = hasSpecifiedLineTradeAgreement;
        HasProductTradePrice = hasProductTradePrice;
        HasPriceChargeAmount = hasPriceChargeAmount;
        MissingLinePricingFields = missingLinePricingFields;
    }

    internal bool HasIncludedSupplyChainTradeLineItem { get; }

    internal bool HasSpecifiedLineTradeAgreement { get; }

    internal bool HasProductTradePrice { get; }

    internal bool HasPriceChargeAmount { get; }

    internal IReadOnlyList<string> MissingLinePricingFields { get; }
}
