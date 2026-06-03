namespace OfficeIMO.Pdf;

internal sealed class PdfCiiAllowanceChargeReasonEvidence {
    internal PdfCiiAllowanceChargeReasonEvidence(
        IReadOnlyList<string> missingAllowanceReasons,
        IReadOnlyList<string> missingChargeReasons) {
        MissingAllowanceReasons = missingAllowanceReasons;
        MissingChargeReasons = missingChargeReasons;
    }

    internal IReadOnlyList<string> MissingAllowanceReasons { get; }

    internal IReadOnlyList<string> MissingChargeReasons { get; }

    internal bool AllAllowanceChargesHaveReason =>
        MissingAllowanceReasons.Count == 0 &&
        MissingChargeReasons.Count == 0;
}
