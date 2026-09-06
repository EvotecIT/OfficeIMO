namespace OfficeIMO.Pdf;

internal sealed class PdfCiiLineAmountConsistencyEvidence {
    internal PdfCiiLineAmountConsistencyEvidence(
        bool hasIncludedSupplyChainTradeLineItem,
        bool hasBilledQuantity,
        bool hasPriceChargeAmount,
        bool hasLineTotalAmount,
        bool allLineAmountsMatch,
        IReadOnlyList<string> mismatchedLineIds,
        string? parseDiagnostic) {
        HasIncludedSupplyChainTradeLineItem = hasIncludedSupplyChainTradeLineItem;
        HasBilledQuantity = hasBilledQuantity;
        HasPriceChargeAmount = hasPriceChargeAmount;
        HasLineTotalAmount = hasLineTotalAmount;
        AllLineAmountsMatch = allLineAmountsMatch;
        MismatchedLineIds = mismatchedLineIds;
        ParseDiagnostic = string.IsNullOrWhiteSpace(parseDiagnostic) ? null : parseDiagnostic!.Trim();
    }

    internal bool HasIncludedSupplyChainTradeLineItem { get; }

    internal bool HasBilledQuantity { get; }

    internal bool HasPriceChargeAmount { get; }

    internal bool HasLineTotalAmount { get; }

    internal bool AllLineAmountsMatch { get; }

    internal IReadOnlyList<string> MismatchedLineIds { get; }

    internal string? ParseDiagnostic { get; }
}
