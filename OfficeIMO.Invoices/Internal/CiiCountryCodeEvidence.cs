namespace OfficeIMO.Invoices.Internal;

internal sealed class CiiCountryCodeEvidence {
    internal CiiCountryCodeEvidence(
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
