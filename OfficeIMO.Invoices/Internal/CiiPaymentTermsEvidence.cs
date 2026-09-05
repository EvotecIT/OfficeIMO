namespace OfficeIMO.Invoices.Internal;

internal sealed class CiiPaymentTermsEvidence {
    internal CiiPaymentTermsEvidence(
        bool hasSpecifiedTradePaymentTerms,
        bool hasDescription,
        bool hasDueDateDateTime) {
        HasSpecifiedTradePaymentTerms = hasSpecifiedTradePaymentTerms;
        HasDescription = hasDescription;
        HasDueDateDateTime = hasDueDateDateTime;
    }

    internal bool HasSpecifiedTradePaymentTerms { get; }

    internal bool HasDescription { get; }

    internal bool HasDueDateDateTime { get; }

    internal bool HasDueDateOrDescription => HasDueDateDateTime || HasDescription;
}
