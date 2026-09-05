namespace OfficeIMO.Invoices.Internal;

internal sealed class CiiPartyTaxRegistrationSchemeEvidence {
    internal CiiPartyTaxRegistrationSchemeEvidence(
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
