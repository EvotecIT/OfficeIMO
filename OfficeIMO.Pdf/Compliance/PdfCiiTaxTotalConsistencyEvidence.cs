namespace OfficeIMO.Pdf;

internal sealed class PdfCiiTaxTotalConsistencyEvidence {
    internal PdfCiiTaxTotalConsistencyEvidence(
        decimal? taxBasisBreakdownSum,
        decimal? taxCalculatedBreakdownSum,
        decimal? taxBasisTotalAmount,
        decimal? taxTotalAmount,
        decimal? notSubjectHeaderBasisAmount,
        decimal? notSubjectLineNetAmountSum,
        decimal? notSubjectAllowanceAmountSum,
        decimal? notSubjectChargeAmountSum,
        IReadOnlyList<string> adjustedBasisMismatches,
        string? parseDiagnostic) {
        TaxBasisBreakdownSum = taxBasisBreakdownSum;
        TaxCalculatedBreakdownSum = taxCalculatedBreakdownSum;
        TaxBasisTotalAmount = taxBasisTotalAmount;
        TaxTotalAmount = taxTotalAmount;
        NotSubjectHeaderBasisAmount = notSubjectHeaderBasisAmount;
        NotSubjectLineNetAmountSum = notSubjectLineNetAmountSum;
        NotSubjectAllowanceAmountSum = notSubjectAllowanceAmountSum;
        NotSubjectChargeAmountSum = notSubjectChargeAmountSum;
        AdjustedBasisMismatches = adjustedBasisMismatches;
        ParseDiagnostic = string.IsNullOrWhiteSpace(parseDiagnostic) ? null : parseDiagnostic!.Trim();
    }

    internal decimal? TaxBasisBreakdownSum { get; }

    internal decimal? TaxCalculatedBreakdownSum { get; }

    internal decimal? TaxBasisTotalAmount { get; }

    internal decimal? TaxTotalAmount { get; }

    internal decimal? NotSubjectHeaderBasisAmount { get; }

    internal decimal? NotSubjectLineNetAmountSum { get; }

    internal decimal? NotSubjectAllowanceAmountSum { get; }

    internal decimal? NotSubjectChargeAmountSum { get; }

    internal IReadOnlyList<string> AdjustedBasisMismatches { get; }

    internal decimal? NotSubjectAdjustedLineNetAmountSum =>
        NotSubjectLineNetAmountSum.HasValue
            ? NotSubjectLineNetAmountSum.Value - (NotSubjectAllowanceAmountSum ?? 0m) + (NotSubjectChargeAmountSum ?? 0m)
            : null;

    internal string? ParseDiagnostic { get; }

    internal bool TaxBasisBreakdownMatchesTotal => AreClose(TaxBasisBreakdownSum, TaxBasisTotalAmount);

    internal bool TaxCalculatedBreakdownMatchesTotal => AreClose(TaxCalculatedBreakdownSum, TaxTotalAmount);

    internal bool NotSubjectHeaderBasisMatchesLineNetAmount => AreClose(NotSubjectHeaderBasisAmount, NotSubjectAdjustedLineNetAmountSum);

    internal bool AdjustedBasisAmountsMatch => AdjustedBasisMismatches.Count == 0;

    private static bool AreClose(decimal? left, decimal? right) =>
        left.HasValue &&
        right.HasValue &&
        System.Math.Abs(left.Value - right.Value) <= 0.01m;
}
