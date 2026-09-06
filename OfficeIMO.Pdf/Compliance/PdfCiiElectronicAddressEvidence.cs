namespace OfficeIMO.Pdf;

internal sealed class PdfCiiElectronicAddressEvidence {
    internal PdfCiiElectronicAddressEvidence(
        bool hasSellerUriUniversalCommunication,
        bool hasSellerUriId,
        bool hasSellerSchemeId,
        IReadOnlyList<string> sellerSchemeIds,
        bool hasBuyerUriUniversalCommunication,
        bool hasBuyerUriId,
        bool hasBuyerSchemeId,
        IReadOnlyList<string> buyerSchemeIds) {
        HasSellerUriUniversalCommunication = hasSellerUriUniversalCommunication;
        HasSellerUriId = hasSellerUriId;
        HasSellerSchemeId = hasSellerSchemeId;
        SellerSchemeIds = sellerSchemeIds;
        HasBuyerUriUniversalCommunication = hasBuyerUriUniversalCommunication;
        HasBuyerUriId = hasBuyerUriId;
        HasBuyerSchemeId = hasBuyerSchemeId;
        BuyerSchemeIds = buyerSchemeIds;
    }

    internal bool HasSellerUriUniversalCommunication { get; }

    internal bool HasSellerUriId { get; }

    internal bool HasSellerSchemeId { get; }

    internal IReadOnlyList<string> SellerSchemeIds { get; }

    internal bool HasBuyerUriUniversalCommunication { get; }

    internal bool HasBuyerUriId { get; }

    internal bool HasBuyerSchemeId { get; }

    internal IReadOnlyList<string> BuyerSchemeIds { get; }
}
