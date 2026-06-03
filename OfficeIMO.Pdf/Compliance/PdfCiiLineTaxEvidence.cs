namespace OfficeIMO.Pdf;

internal sealed class PdfCiiLineTaxEvidence {
    internal PdfCiiLineTaxEvidence(
        bool hasIncludedSupplyChainTradeLineItem,
        bool hasSpecifiedLineTradeSettlement,
        bool hasApplicableTradeTax,
        bool hasTypeCode,
        bool hasCategoryCode,
        bool hasRateApplicablePercent,
        bool hasRateRequirementCoverage,
        IReadOnlyList<string> typeCodes,
        IReadOnlyList<string> missingRateCategoryCodes,
        IReadOnlyList<string> forbiddenRateCategoryCodes) {
        HasIncludedSupplyChainTradeLineItem = hasIncludedSupplyChainTradeLineItem;
        HasSpecifiedLineTradeSettlement = hasSpecifiedLineTradeSettlement;
        HasApplicableTradeTax = hasApplicableTradeTax;
        HasTypeCode = hasTypeCode;
        HasCategoryCode = hasCategoryCode;
        HasRateApplicablePercent = hasRateApplicablePercent;
        HasRateRequirementCoverage = hasRateRequirementCoverage;
        TypeCodes = typeCodes;
        MissingRateCategoryCodes = missingRateCategoryCodes;
        ForbiddenRateCategoryCodes = forbiddenRateCategoryCodes;
    }

    internal bool HasIncludedSupplyChainTradeLineItem { get; }

    internal bool HasSpecifiedLineTradeSettlement { get; }

    internal bool HasApplicableTradeTax { get; }

    internal bool HasTypeCode { get; }

    internal bool HasCategoryCode { get; }

    internal bool HasRateApplicablePercent { get; }

    internal bool HasRateRequirementCoverage { get; }

    internal IReadOnlyList<string> TypeCodes { get; }

    internal IReadOnlyList<string> MissingRateCategoryCodes { get; }

    internal IReadOnlyList<string> ForbiddenRateCategoryCodes { get; }
}
