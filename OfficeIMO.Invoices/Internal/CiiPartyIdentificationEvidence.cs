namespace OfficeIMO.Invoices.Internal;

internal sealed class CiiPartyIdentificationEvidence {
    internal CiiPartyIdentificationEvidence(
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
