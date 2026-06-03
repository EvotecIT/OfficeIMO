namespace OfficeIMO.Pdf;

internal sealed class PdfCiiTaxBreakdownEvidence {
    internal PdfCiiTaxBreakdownEvidence(
        bool hasApplicableTradeTax,
        bool hasTypeCode,
        bool hasCategoryCode,
        bool hasRateApplicablePercent,
        bool hasBasisAmount,
        bool hasCalculatedAmount,
        IReadOnlyList<string> typeCodes,
        IReadOnlyList<string> missingTypeCodeBreakdowns) {
        HasApplicableTradeTax = hasApplicableTradeTax;
        HasTypeCode = hasTypeCode;
        HasCategoryCode = hasCategoryCode;
        HasRateApplicablePercent = hasRateApplicablePercent;
        HasBasisAmount = hasBasisAmount;
        HasCalculatedAmount = hasCalculatedAmount;
        TypeCodes = typeCodes;
        MissingTypeCodeBreakdowns = missingTypeCodeBreakdowns;
    }

    internal bool HasApplicableTradeTax { get; }

    internal bool HasTypeCode { get; }

    internal bool HasCategoryCode { get; }

    internal bool HasRateApplicablePercent { get; }

    internal bool HasBasisAmount { get; }

    internal bool HasCalculatedAmount { get; }

    internal IReadOnlyList<string> TypeCodes { get; }

    internal IReadOnlyList<string> MissingTypeCodeBreakdowns { get; }
}
