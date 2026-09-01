namespace OfficeIMO.Pdf;

/// <summary>Preview of text, image placements, and annotations that intersect requested redaction rectangles.</summary>
public sealed class PdfRedactionPlan {
    internal PdfRedactionPlan(
        PdfDocumentPreflight preflight,
        IReadOnlyList<PdfRedactionArea> areas,
        IReadOnlyList<PdfRedactionMatch> matches,
        IReadOnlyList<PdfDiagnosticFinding> findings,
        IReadOnlyList<string>? searchCriteria,
        string sourceSha256) {
        Preflight = preflight;
        Areas = areas;
        Matches = matches;
        Findings = findings;
        SearchCriteria = searchCriteria ?? Array.Empty<string>();
        SourceSha256 = sourceSha256;
    }

    /// <summary>Preflight result used while creating the plan.</summary>
    public PdfDocumentPreflight Preflight { get; }

    /// <summary>Requested redaction areas.</summary>
    public IReadOnlyList<PdfRedactionArea> Areas { get; }

    /// <summary>Text blocks, image placements, and annotations intersecting the requested areas.</summary>
    public IReadOnlyList<PdfRedactionMatch> Matches { get; }

    /// <summary>Diagnostics and warnings for the plan.</summary>
    public IReadOnlyList<PdfDiagnosticFinding> Findings { get; }

    /// <summary>Stable descriptions of literal, regex, logical-kind, or form-field criteria used to derive the areas.</summary>
    public IReadOnlyList<string> SearchCriteria { get; }

    /// <summary>SHA-256 fingerprint of the exact PDF bytes inspected while creating this plan.</summary>
    public string SourceSha256 { get; }

    /// <summary>True when the source was inspectable and the plan contains no blocking findings.</summary>
    public bool IsReviewable =>
        Preflight.CanReadLogicalObjects &&
        Findings.All(static finding => finding.Severity != PdfDiagnosticSeverity.Error);

    /// <summary>True when the plan areas were derived from explicit search criteria.</summary>
    public bool IsSearchDriven => SearchCriteria.Count > 0;

    /// <summary>True when at least one match was found.</summary>
    public bool HasMatches => Matches.Count > 0;

    internal bool MatchesSource(byte[] pdf) =>
        string.Equals(SourceSha256, ComputeSourceSha256(pdf), StringComparison.Ordinal);

    internal static string ComputeSourceSha256(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));
#if NET6_0_OR_GREATER
        return Convert.ToBase64String(System.Security.Cryptography.SHA256.HashData(pdf));
#else
        using var sha256 = System.Security.Cryptography.SHA256.Create();
        return Convert.ToBase64String(sha256.ComputeHash(pdf));
#endif
    }
}
