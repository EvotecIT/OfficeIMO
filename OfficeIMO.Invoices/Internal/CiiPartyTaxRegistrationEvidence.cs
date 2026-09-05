namespace OfficeIMO.Invoices.Internal;

internal sealed class CiiPartyTaxRegistrationEvidence {
    internal CiiPartyTaxRegistrationEvidence(
        bool hasSellerTaxRegistrationId,
        bool hasBuyerTaxRegistrationId) {
        HasSellerTaxRegistrationId = hasSellerTaxRegistrationId;
        HasBuyerTaxRegistrationId = hasBuyerTaxRegistrationId;
    }

    internal bool HasSellerTaxRegistrationId { get; }

    internal bool HasBuyerTaxRegistrationId { get; }
}
