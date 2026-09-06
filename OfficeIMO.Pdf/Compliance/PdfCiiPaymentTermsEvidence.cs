namespace OfficeIMO.Pdf;

internal sealed class PdfCiiPaymentTermsEvidence {
    internal PdfCiiPaymentTermsEvidence(
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
