namespace OfficeIMO.Pdf;

internal sealed class PdfCiiTaxExemptionReasonEvidence {
    internal PdfCiiTaxExemptionReasonEvidence(
        bool hasApplicableTradeTax,
        bool allRequiredCategoriesHaveReason,
        IReadOnlyList<string> missingReasonCategories,
        bool allForbiddenCategoriesOmitReason,
        IReadOnlyList<string> forbiddenReasonCategories,
        bool allNotSubjectReasonCodesAreCanonical,
        IReadOnlyList<string> invalidNotSubjectReasonCodes) {
        HasApplicableTradeTax = hasApplicableTradeTax;
        AllRequiredCategoriesHaveReason = allRequiredCategoriesHaveReason;
        MissingReasonCategories = missingReasonCategories;
        AllForbiddenCategoriesOmitReason = allForbiddenCategoriesOmitReason;
        ForbiddenReasonCategories = forbiddenReasonCategories;
        AllNotSubjectReasonCodesAreCanonical = allNotSubjectReasonCodesAreCanonical;
        InvalidNotSubjectReasonCodes = invalidNotSubjectReasonCodes;
    }

    internal bool HasApplicableTradeTax { get; }

    internal bool AllRequiredCategoriesHaveReason { get; }

    internal IReadOnlyList<string> MissingReasonCategories { get; }

    internal bool AllForbiddenCategoriesOmitReason { get; }

    internal IReadOnlyList<string> ForbiddenReasonCategories { get; }

    internal bool AllNotSubjectReasonCodesAreCanonical { get; }

    internal IReadOnlyList<string> InvalidNotSubjectReasonCodes { get; }
}
