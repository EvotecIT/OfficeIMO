namespace OfficeIMO.Pdf;

internal sealed class PdfCiiTaxPartyIdentifierEvidence {
    internal PdfCiiTaxPartyIdentifierEvidence(
        bool hasApplicableTradeTax,
        bool hasSellerTaxRegistrationId,
        bool hasSellerVatRegistrationId,
        bool hasBuyerTaxRegistrationId,
        bool hasBuyerVatRegistrationId,
        IReadOnlyList<string> missingSellerIdentifierCategories,
        IReadOnlyList<string> missingBuyerIdentifierCategories,
        IReadOnlyList<string> forbiddenSellerVatIdentifierCategories,
        IReadOnlyList<string> forbiddenBuyerVatIdentifierCategories) {
        HasApplicableTradeTax = hasApplicableTradeTax;
        HasSellerTaxRegistrationId = hasSellerTaxRegistrationId;
        HasSellerVatRegistrationId = hasSellerVatRegistrationId;
        HasBuyerTaxRegistrationId = hasBuyerTaxRegistrationId;
        HasBuyerVatRegistrationId = hasBuyerVatRegistrationId;
        MissingSellerIdentifierCategories = missingSellerIdentifierCategories;
        MissingBuyerIdentifierCategories = missingBuyerIdentifierCategories;
        ForbiddenSellerVatIdentifierCategories = forbiddenSellerVatIdentifierCategories;
        ForbiddenBuyerVatIdentifierCategories = forbiddenBuyerVatIdentifierCategories;
    }

    internal bool HasApplicableTradeTax { get; }

    internal bool HasSellerTaxRegistrationId { get; }

    internal bool HasSellerVatRegistrationId { get; }

    internal bool HasBuyerTaxRegistrationId { get; }

    internal bool HasBuyerVatRegistrationId { get; }

    internal IReadOnlyList<string> MissingSellerIdentifierCategories { get; }

    internal IReadOnlyList<string> MissingBuyerIdentifierCategories { get; }

    internal IReadOnlyList<string> ForbiddenSellerVatIdentifierCategories { get; }

    internal IReadOnlyList<string> ForbiddenBuyerVatIdentifierCategories { get; }
}
