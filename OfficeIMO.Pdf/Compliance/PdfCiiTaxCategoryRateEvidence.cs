namespace OfficeIMO.Pdf;

internal sealed class PdfCiiTaxCategoryRateEvidence {
    internal PdfCiiTaxCategoryRateEvidence(
        bool hasApplicableTradeTax,
        bool hasTaxCategoryRate,
        bool hasRateRequirementCoverage,
        bool allZeroRatedCategoriesUseZeroRate,
        IReadOnlyList<string> nonZeroRatedCategoryRates,
        IReadOnlyList<string> missingRateCategoryCodes,
        IReadOnlyList<string> forbiddenRateCategoryCodes,
        string? parseDiagnostic) {
        HasApplicableTradeTax = hasApplicableTradeTax;
        HasTaxCategoryRate = hasTaxCategoryRate;
        HasRateRequirementCoverage = hasRateRequirementCoverage;
        AllZeroRatedCategoriesUseZeroRate = allZeroRatedCategoriesUseZeroRate;
        NonZeroRatedCategoryRates = nonZeroRatedCategoryRates;
        MissingRateCategoryCodes = missingRateCategoryCodes;
        ForbiddenRateCategoryCodes = forbiddenRateCategoryCodes;
        ParseDiagnostic = parseDiagnostic;
    }

    internal bool HasApplicableTradeTax { get; }

    internal bool HasTaxCategoryRate { get; }

    internal bool HasRateRequirementCoverage { get; }

    internal bool AllZeroRatedCategoriesUseZeroRate { get; }

    internal IReadOnlyList<string> NonZeroRatedCategoryRates { get; }

    internal IReadOnlyList<string> MissingRateCategoryCodes { get; }

    internal IReadOnlyList<string> ForbiddenRateCategoryCodes { get; }

    internal string? ParseDiagnostic { get; }
}
