namespace OfficeIMO.Pdf;

internal sealed class PdfCiiTaxCategoryConsistencyEvidence {
    internal PdfCiiTaxCategoryConsistencyEvidence(
        bool hasHeaderApplicableTradeTax,
        bool hasLineApplicableTradeTax,
        bool hasHeaderTaxCategoryRate,
        bool hasLineTaxCategoryRate,
        bool allLineTaxCategoryRatesMatchHeaderBreakdown,
        IReadOnlyList<string> unmatchedLineTaxCategoryRates,
        bool allAllowanceChargeTaxCategoryRatesMatchHeaderBreakdown,
        IReadOnlyList<string> unmatchedAllowanceChargeTaxCategoryRates) {
        HasHeaderApplicableTradeTax = hasHeaderApplicableTradeTax;
        HasLineApplicableTradeTax = hasLineApplicableTradeTax;
        HasHeaderTaxCategoryRate = hasHeaderTaxCategoryRate;
        HasLineTaxCategoryRate = hasLineTaxCategoryRate;
        AllLineTaxCategoryRatesMatchHeaderBreakdown = allLineTaxCategoryRatesMatchHeaderBreakdown;
        UnmatchedLineTaxCategoryRates = unmatchedLineTaxCategoryRates;
        AllAllowanceChargeTaxCategoryRatesMatchHeaderBreakdown = allAllowanceChargeTaxCategoryRatesMatchHeaderBreakdown;
        UnmatchedAllowanceChargeTaxCategoryRates = unmatchedAllowanceChargeTaxCategoryRates;
    }

    internal bool HasHeaderApplicableTradeTax { get; }

    internal bool HasLineApplicableTradeTax { get; }

    internal bool HasHeaderTaxCategoryRate { get; }

    internal bool HasLineTaxCategoryRate { get; }

    internal bool AllLineTaxCategoryRatesMatchHeaderBreakdown { get; }

    internal IReadOnlyList<string> UnmatchedLineTaxCategoryRates { get; }

    internal bool AllAllowanceChargeTaxCategoryRatesMatchHeaderBreakdown { get; }

    internal IReadOnlyList<string> UnmatchedAllowanceChargeTaxCategoryRates { get; }
}
