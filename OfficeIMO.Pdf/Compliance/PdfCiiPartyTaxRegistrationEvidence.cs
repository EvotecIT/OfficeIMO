namespace OfficeIMO.Pdf;

internal sealed class PdfCiiPartyTaxRegistrationEvidence {
    internal PdfCiiPartyTaxRegistrationEvidence(
        bool hasSellerTaxRegistrationId,
        bool hasBuyerTaxRegistrationId) {
        HasSellerTaxRegistrationId = hasSellerTaxRegistrationId;
        HasBuyerTaxRegistrationId = hasBuyerTaxRegistrationId;
    }

    internal bool HasSellerTaxRegistrationId { get; }

    internal bool HasBuyerTaxRegistrationId { get; }
}
