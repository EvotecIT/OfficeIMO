namespace OfficeIMO.Invoices.Internal;

internal sealed class CiiUnitCodeEvidence {
    internal CiiUnitCodeEvidence(
        bool hasIncludedSupplyChainTradeLineItem,
        bool hasBilledQuantity,
        bool hasBilledQuantityUnitCode,
        IReadOnlyList<string> unitCodes) {
        HasIncludedSupplyChainTradeLineItem = hasIncludedSupplyChainTradeLineItem;
        HasBilledQuantity = hasBilledQuantity;
        HasBilledQuantityUnitCode = hasBilledQuantityUnitCode;
        UnitCodes = unitCodes;
    }

    internal bool HasIncludedSupplyChainTradeLineItem { get; }

    internal bool HasBilledQuantity { get; }

    internal bool HasBilledQuantityUnitCode { get; }

    internal IReadOnlyList<string> UnitCodes { get; }
}
