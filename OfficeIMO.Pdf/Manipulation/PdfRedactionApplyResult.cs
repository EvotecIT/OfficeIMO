namespace OfficeIMO.Pdf;

/// <summary>Verification state for one item from a reviewed redaction plan.</summary>
public enum PdfRedactionEvidenceStatus {
    /// <summary>The rewritten PDF was fully inspected and no matching content remained in the reviewed area.</summary>
    VerifiedAbsent = 0,
    /// <summary>Matching content remains in the reviewed area after the rewrite.</summary>
    Residual = 1,
    /// <summary>The item could not be proven absent because another verification check failed or was blocked.</summary>
    Inconclusive = 2
}

/// <summary>Post-rewrite evidence for one item from a reviewed redaction plan.</summary>
public sealed class PdfRedactionEvidenceItem {
    internal PdfRedactionEvidenceItem(
        PdfRedactionMatch reviewedMatch,
        PdfRedactionEvidenceStatus status,
        IReadOnlyList<PdfRedactionMatch> residualMatches) {
        ReviewedMatch = reviewedMatch;
        Status = status;
        ResidualMatches = residualMatches;
    }

    /// <summary>Content item approved for removal in the source-bound plan.</summary>
    public PdfRedactionMatch ReviewedMatch { get; }

    /// <summary>Outcome established by inspecting the rewritten artifact.</summary>
    public PdfRedactionEvidenceStatus Status { get; }

    /// <summary>Post-rewrite matches of the same kind in the same reviewed area.</summary>
    public IReadOnlyList<PdfRedactionMatch> ResidualMatches { get; }
}

/// <summary>Source-bound, actual-versus-planned evidence for a redaction rewrite.</summary>
public sealed class PdfRedactionEvidenceReport {
    internal PdfRedactionEvidenceReport(
        PdfRedactionPlan reviewedPlan,
        string outputSha256,
        IReadOnlyList<PdfRedactionMatch> residualMatches,
        PdfRedactionVerificationReport verification) {
        ReviewedPlan = reviewedPlan;
        OutputSha256 = outputSha256;
        ResidualMatches = residualMatches;
        Verification = verification;
        AffectedPageNumbers = reviewedPlan.Areas
            .Select(static area => area.PageNumber)
            .Distinct()
            .OrderBy(static pageNumber => pageNumber)
            .ToArray();
        Items = BuildItems(reviewedPlan.Matches, residualMatches, verification.IsVerified);
    }

    /// <summary>Reviewed plan bound to the exact source bytes that were rewritten.</summary>
    public PdfRedactionPlan ReviewedPlan { get; }

    /// <summary>SHA-256 fingerprint of the reviewed source PDF.</summary>
    public string SourceSha256 => ReviewedPlan.SourceSha256;

    /// <summary>SHA-256 fingerprint of the rewritten PDF.</summary>
    public string OutputSha256 { get; }

    /// <summary>One-based page numbers affected by the reviewed areas, suitable for targeted preview rendering.</summary>
    public IReadOnlyList<int> AffectedPageNumbers { get; }

    /// <summary>Per-planned-item outcomes established from the rewritten artifact.</summary>
    public IReadOnlyList<PdfRedactionEvidenceItem> Items { get; }

    /// <summary>Content still intersecting reviewed areas in the rewritten PDF.</summary>
    public IReadOnlyList<PdfRedactionMatch> ResidualMatches { get; }

    /// <summary>Marker, stream, rendering, external-validator, page-identity, and residual-content checks.</summary>
    public PdfRedactionVerificationReport Verification { get; }

    /// <summary>True when every configured verification check passed and all planned items were proven absent.</summary>
    public bool IsVerified => Verification.IsVerified && Items.All(static item => item.Status == PdfRedactionEvidenceStatus.VerifiedAbsent);

    /// <summary>Number of reviewed items proven absent from their reviewed areas.</summary>
    public int VerifiedAbsentCount => Items.Count(static item => item.Status == PdfRedactionEvidenceStatus.VerifiedAbsent);

    /// <summary>Number of reviewed items with residual content of the same kind in their reviewed areas.</summary>
    public int ResidualCount => Items.Count(static item => item.Status == PdfRedactionEvidenceStatus.Residual);

    /// <summary>Number of reviewed items whose absence could not be proven.</summary>
    public int InconclusiveCount => Items.Count(static item => item.Status == PdfRedactionEvidenceStatus.Inconclusive);

    /// <summary>Human-readable evidence summary suitable for logs and thin clients.</summary>
    public string Summary => IsVerified
        ? $"Verified removal of {VerifiedAbsentCount} reviewed PDF content item(s) across {AffectedPageNumbers.Count} page(s)."
        : $"PDF redaction evidence is not complete: {ResidualCount} residual and {InconclusiveCount} inconclusive reviewed item(s). {Verification.Summary}";

    private static PdfRedactionEvidenceItem[] BuildItems(
        IReadOnlyList<PdfRedactionMatch> reviewedMatches,
        IReadOnlyList<PdfRedactionMatch> residualMatches,
        bool verificationPassed) {
        var items = new PdfRedactionEvidenceItem[reviewedMatches.Count];
        for (int i = 0; i < reviewedMatches.Count; i++) {
            PdfRedactionMatch reviewed = reviewedMatches[i];
            PdfRedactionMatch[] residual = residualMatches
                .Where(candidate => candidate.Kind == reviewed.Kind &&
                    candidate.PageNumber == reviewed.PageNumber &&
                    SameArea(candidate.Area, reviewed.Area))
                .ToArray();
            PdfRedactionEvidenceStatus status = residual.Length > 0
                ? PdfRedactionEvidenceStatus.Residual
                : verificationPassed
                    ? PdfRedactionEvidenceStatus.VerifiedAbsent
                    : PdfRedactionEvidenceStatus.Inconclusive;
            items[i] = new PdfRedactionEvidenceItem(reviewed, status, residual);
        }

        return items;
    }

    private static bool SameArea(PdfRedactionArea left, PdfRedactionArea right) =>
        left.PageNumber == right.PageNumber &&
        Math.Abs(left.X - right.X) <= 0.001D &&
        Math.Abs(left.Y - right.Y) <= 0.001D &&
        Math.Abs(left.Width - right.Width) <= 0.001D &&
        Math.Abs(left.Height - right.Height) <= 0.001D;
}

/// <summary>Rewritten PDF bytes together with the selected mutation path and redaction evidence.</summary>
public sealed class PdfRedactionApplyResult {
    private readonly byte[] _pdf;
    private readonly PdfLoadOptions? _readOptions;

    internal PdfRedactionApplyResult(
        byte[] pdf,
        PdfMutationPlan mutationPlan,
        PdfRedactionEvidenceReport evidence,
        PdfLoadOptions? readOptions) {
        _pdf = (byte[])pdf.Clone();
        MutationPlan = mutationPlan;
        Evidence = evidence;
        _readOptions = readOptions;
    }

    /// <summary>Rewritten PDF bytes.</summary>
    public byte[] Pdf => (byte[])_pdf.Clone();

    /// <summary>Full-rewrite decision used for the redaction mutation.</summary>
    public PdfMutationPlan MutationPlan { get; }

    /// <summary>Source-bound actual-versus-planned redaction evidence.</summary>
    public PdfRedactionEvidenceReport Evidence { get; }

    /// <summary>True when the rewritten artifact passed every configured evidence check.</summary>
    public bool IsVerified => Evidence.IsVerified;

    /// <summary>Opens the rewritten bytes through the normal fluent document API.</summary>
    public PdfDocument ToDocument() => PdfDocument.Load(_pdf, _readOptions);

    /// <summary>Throws when the rewritten artifact did not pass every configured evidence check.</summary>
    public PdfRedactionApplyResult ThrowIfUnverified() {
        if (!IsVerified) throw new InvalidOperationException(Evidence.Summary);
        return this;
    }
}
