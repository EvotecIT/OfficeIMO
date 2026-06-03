namespace OfficeIMO.Pdf;

internal sealed class PdfCiiLineItemEvidence {
    internal PdfCiiLineItemEvidence(
        bool hasIncludedSupplyChainTradeLineItem,
        bool hasLineId,
        bool hasProductName,
        bool hasBilledQuantity,
        bool hasBilledQuantityUnitCode,
        bool hasLineTotalAmount) {
        HasIncludedSupplyChainTradeLineItem = hasIncludedSupplyChainTradeLineItem;
        HasLineId = hasLineId;
        HasProductName = hasProductName;
        HasBilledQuantity = hasBilledQuantity;
        HasBilledQuantityUnitCode = hasBilledQuantityUnitCode;
        HasLineTotalAmount = hasLineTotalAmount;
    }

    internal bool HasIncludedSupplyChainTradeLineItem { get; }

    internal bool HasLineId { get; }

    internal bool HasProductName { get; }

    internal bool HasBilledQuantity { get; }

    internal bool HasBilledQuantityUnitCode { get; }

    internal bool HasLineTotalAmount { get; }
}
