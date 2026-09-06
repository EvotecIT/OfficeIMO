namespace OfficeIMO.Pdf;

internal sealed class PdfCiiPartyIdentificationEvidence {
    internal PdfCiiPartyIdentificationEvidence(
        bool hasSellerName,
        bool hasSellerCountryId,
        bool hasBuyerName,
        bool hasBuyerCountryId) {
        HasSellerName = hasSellerName;
        HasSellerCountryId = hasSellerCountryId;
        HasBuyerName = hasBuyerName;
        HasBuyerCountryId = hasBuyerCountryId;
    }

    internal bool HasSellerName { get; }

    internal bool HasSellerCountryId { get; }

    internal bool HasBuyerName { get; }

    internal bool HasBuyerCountryId { get; }
}
