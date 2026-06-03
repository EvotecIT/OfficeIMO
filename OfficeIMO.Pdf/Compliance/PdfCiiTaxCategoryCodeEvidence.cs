namespace OfficeIMO.Pdf;

internal sealed class PdfCiiTaxCategoryCodeEvidence {
    internal PdfCiiTaxCategoryCodeEvidence(
        bool hasApplicableTradeTax,
        bool hasCategoryCode,
        IReadOnlyList<string> categoryCodes,
        int headerNotSubjectTaxBreakdownCount,
        IReadOnlyList<string> nonNotSubjectHeaderTaxBreakdownCategories,
        bool hasLineNotSubjectTaxCategory,
        bool hasAllowanceChargeNotSubjectTaxCategory) {
        HasApplicableTradeTax = hasApplicableTradeTax;
        HasCategoryCode = hasCategoryCode;
        CategoryCodes = categoryCodes;
        HeaderNotSubjectTaxBreakdownCount = headerNotSubjectTaxBreakdownCount;
        NonNotSubjectHeaderTaxBreakdownCategories = nonNotSubjectHeaderTaxBreakdownCategories;
        HasLineNotSubjectTaxCategory = hasLineNotSubjectTaxCategory;
        HasAllowanceChargeNotSubjectTaxCategory = hasAllowanceChargeNotSubjectTaxCategory;
    }

    internal bool HasApplicableTradeTax { get; }

    internal bool HasCategoryCode { get; }

    internal IReadOnlyList<string> CategoryCodes { get; }

    internal int HeaderNotSubjectTaxBreakdownCount { get; }

    internal IReadOnlyList<string> NonNotSubjectHeaderTaxBreakdownCategories { get; }

    internal bool HasLineNotSubjectTaxCategory { get; }

    internal bool HasAllowanceChargeNotSubjectTaxCategory { get; }
}
