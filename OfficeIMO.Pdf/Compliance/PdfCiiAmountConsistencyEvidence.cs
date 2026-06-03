namespace OfficeIMO.Pdf;

internal sealed class PdfCiiAmountConsistencyEvidence {
    internal PdfCiiAmountConsistencyEvidence(
        decimal? lineTotalAmountSum,
        decimal? allowanceTotalAmount,
        decimal? chargeTotalAmount,
        decimal? documentLevelAllowanceAmountSum,
        decimal? documentLevelChargeAmountSum,
        decimal? taxBasisTotalAmount,
        decimal? taxTotalAmount,
        decimal? grandTotalAmount,
        decimal? duePayableAmount,
        decimal? paidAmount,
        decimal? roundingAmount,
        string? parseDiagnostic) {
        LineTotalAmountSum = lineTotalAmountSum;
        AllowanceTotalAmount = allowanceTotalAmount;
        ChargeTotalAmount = chargeTotalAmount;
        DocumentLevelAllowanceAmountSum = documentLevelAllowanceAmountSum;
        DocumentLevelChargeAmountSum = documentLevelChargeAmountSum;
        TaxBasisTotalAmount = taxBasisTotalAmount;
        TaxTotalAmount = taxTotalAmount;
        GrandTotalAmount = grandTotalAmount;
        DuePayableAmount = duePayableAmount;
        PaidAmount = paidAmount;
        RoundingAmount = roundingAmount;
        ParseDiagnostic = string.IsNullOrWhiteSpace(parseDiagnostic) ? null : parseDiagnostic!.Trim();
    }

    internal decimal? LineTotalAmountSum { get; }

    internal decimal? AllowanceTotalAmount { get; }

    internal decimal? ChargeTotalAmount { get; }

    internal decimal? DocumentLevelAllowanceAmountSum { get; }

    internal decimal? DocumentLevelChargeAmountSum { get; }

    internal decimal? TaxBasisTotalAmount { get; }

    internal decimal? TaxTotalAmount { get; }

    internal decimal? GrandTotalAmount { get; }

    internal decimal? DuePayableAmount { get; }

    internal decimal? PaidAmount { get; }

    internal decimal? RoundingAmount { get; }

    internal decimal? GrandOrDuePayableAmount => GrandTotalAmount ?? DuePayableAmount;

    internal string? ParseDiagnostic { get; }

    internal bool HasRequiredAmounts =>
        LineTotalAmountSum.HasValue &&
        TaxBasisTotalAmount.HasValue &&
        TaxTotalAmount.HasValue &&
        GrandOrDuePayableAmount.HasValue;

    internal bool LineTotalMatchesTaxBasis =>
        LineTotalAmountSum.HasValue &&
        TaxBasisTotalAmount.HasValue &&
        AreClose(LineTotalAmountSum.Value - (AllowanceTotalAmount ?? 0m) + (ChargeTotalAmount ?? 0m), TaxBasisTotalAmount.Value);

    internal bool AllowanceTotalMatchesDocumentLevelAllowances =>
        OptionalTotalMatches(AllowanceTotalAmount, DocumentLevelAllowanceAmountSum);

    internal bool ChargeTotalMatchesDocumentLevelCharges =>
        OptionalTotalMatches(ChargeTotalAmount, DocumentLevelChargeAmountSum);

    internal bool GrandTotalMatchesBasisPlusTax =>
        TaxBasisTotalAmount.HasValue &&
        TaxTotalAmount.HasValue &&
        (GrandTotalAmount.HasValue || DuePayableAmount.HasValue) &&
        AreClose(TaxBasisTotalAmount.Value + TaxTotalAmount.Value, GrandTotalAmount ?? CalculateGrossPayableAmount());

    internal bool DuePayableMatchesGrandTotal =>
        !DuePayableAmount.HasValue ||
        AreClose(CalculateGrossPayableAmount() - (PaidAmount ?? 0m) + (RoundingAmount ?? 0m), DuePayableAmount.Value);

    private decimal? CalculateGrossPayableAmount() {
        if (GrandTotalAmount.HasValue) {
            return GrandTotalAmount.Value;
        }

        if (TaxBasisTotalAmount.HasValue && TaxTotalAmount.HasValue) {
            return TaxBasisTotalAmount.Value + TaxTotalAmount.Value;
        }

        return null;
    }

    private static bool AreClose(decimal? left, decimal? right) =>
        left.HasValue &&
        right.HasValue &&
        System.Math.Abs(left.Value - right.Value) <= 0.01m;

    private static bool OptionalTotalMatches(decimal? total, decimal? componentSum) {
        if (!total.HasValue && !componentSum.HasValue) {
            return true;
        }

        if (total.HasValue && !componentSum.HasValue) {
            return AreClose(total.Value, 0m);
        }

        if (!total.HasValue && componentSum.HasValue) {
            return AreClose(decimal.Round(componentSum.Value, 2, MidpointRounding.AwayFromZero), 0m);
        }

        return AreClose(total.GetValueOrDefault(), decimal.Round(componentSum.GetValueOrDefault(), 2, MidpointRounding.AwayFromZero));
    }
}
