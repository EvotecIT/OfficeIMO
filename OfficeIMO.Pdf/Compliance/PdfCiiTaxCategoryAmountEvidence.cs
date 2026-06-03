namespace OfficeIMO.Pdf;

internal sealed class PdfCiiTaxCategoryAmountEvidence {
    internal PdfCiiTaxCategoryAmountEvidence(
        bool hasApplicableTradeTax,
        bool hasTaxCategoryAmount,
        bool allZeroRatedCategoriesUseZeroAmount,
        IReadOnlyList<string> nonZeroRatedCategoryAmounts,
        bool allStandardRatedCategoryAmountsMatchRate,
        IReadOnlyList<string> mismatchedStandardRatedCategoryAmounts,
        string? parseDiagnostic) {
        HasApplicableTradeTax = hasApplicableTradeTax;
        HasTaxCategoryAmount = hasTaxCategoryAmount;
        AllZeroRatedCategoriesUseZeroAmount = allZeroRatedCategoriesUseZeroAmount;
        NonZeroRatedCategoryAmounts = nonZeroRatedCategoryAmounts;
        AllStandardRatedCategoryAmountsMatchRate = allStandardRatedCategoryAmountsMatchRate;
        MismatchedStandardRatedCategoryAmounts = mismatchedStandardRatedCategoryAmounts;
        ParseDiagnostic = parseDiagnostic;
    }

    internal bool HasApplicableTradeTax { get; }

    internal bool HasTaxCategoryAmount { get; }

    internal bool AllZeroRatedCategoriesUseZeroAmount { get; }

    internal IReadOnlyList<string> NonZeroRatedCategoryAmounts { get; }

    internal bool AllStandardRatedCategoryAmountsMatchRate { get; }

    internal IReadOnlyList<string> MismatchedStandardRatedCategoryAmounts { get; }

    internal string? ParseDiagnostic { get; }
}
