namespace OfficeIMO.Invoices.Internal;

internal sealed class CiiLinePricingEvidence {
    internal CiiLinePricingEvidence(
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
