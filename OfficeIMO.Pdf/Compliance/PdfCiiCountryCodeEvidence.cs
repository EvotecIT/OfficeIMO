namespace OfficeIMO.Pdf;

internal sealed class PdfCiiCountryCodeEvidence {
    internal PdfCiiCountryCodeEvidence(
        bool hasSellerCountryId,
        bool hasBuyerCountryId,
        string? sellerCountryId,
        string? buyerCountryId) {
        HasSellerCountryId = hasSellerCountryId;
        HasBuyerCountryId = hasBuyerCountryId;
        SellerCountryId = sellerCountryId;
        BuyerCountryId = buyerCountryId;
    }

    internal bool HasSellerCountryId { get; }

    internal bool HasBuyerCountryId { get; }

    internal string? SellerCountryId { get; }

    internal string? BuyerCountryId { get; }
}
