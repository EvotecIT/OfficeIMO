namespace OfficeIMO.Invoices.Internal;

internal sealed class CiiDateFormatEvidence {
    internal CiiDateFormatEvidence(
        bool hasIssueDateTime,
        bool issueDateTimeIsParseable,
        bool hasPaymentDueDateTime,
        bool paymentDueDateTimeIsParseable,
        IReadOnlyList<string> invalidDateFields) {
        HasIssueDateTime = hasIssueDateTime;
        IssueDateTimeIsParseable = issueDateTimeIsParseable;
        HasPaymentDueDateTime = hasPaymentDueDateTime;
        PaymentDueDateTimeIsParseable = paymentDueDateTimeIsParseable;
        InvalidDateFields = invalidDateFields;
    }

    internal bool HasIssueDateTime { get; }

    internal bool IssueDateTimeIsParseable { get; }

    internal bool HasPaymentDueDateTime { get; }

    internal bool PaymentDueDateTimeIsParseable { get; }

    internal IReadOnlyList<string> InvalidDateFields { get; }

    internal bool AllKnownDatesAreParseable =>
        (!HasIssueDateTime || IssueDateTimeIsParseable) &&
        (!HasPaymentDueDateTime || PaymentDueDateTimeIsParseable);
}
