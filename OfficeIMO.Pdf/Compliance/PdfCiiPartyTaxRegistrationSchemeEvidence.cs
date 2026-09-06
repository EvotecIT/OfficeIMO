namespace OfficeIMO.Pdf;

internal sealed class PdfCiiPartyTaxRegistrationSchemeEvidence {
    internal PdfCiiPartyTaxRegistrationSchemeEvidence(
        bool hasSellerTaxRegistrationId,
        bool hasSellerTaxRegistrationSchemeId,
        bool hasBuyerTaxRegistrationId,
        bool hasBuyerTaxRegistrationSchemeId) {
        HasSellerTaxRegistrationId = hasSellerTaxRegistrationId;
        HasSellerTaxRegistrationSchemeId = hasSellerTaxRegistrationSchemeId;
        HasBuyerTaxRegistrationId = hasBuyerTaxRegistrationId;
        HasBuyerTaxRegistrationSchemeId = hasBuyerTaxRegistrationSchemeId;
    }

    internal bool HasSellerTaxRegistrationId { get; }

    internal bool HasSellerTaxRegistrationSchemeId { get; }

    internal bool HasBuyerTaxRegistrationId { get; }

    internal bool HasBuyerTaxRegistrationSchemeId { get; }
}
