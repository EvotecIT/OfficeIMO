namespace OfficeIMO.Pdf;

/// <summary>Preview of text, image placements, and annotations that intersect requested redaction rectangles.</summary>
public sealed class PdfRedactionPlan {
    internal PdfRedactionPlan(
        PdfDocumentPreflight preflight,
        IReadOnlyList<PdfRedactionArea> areas,
        IReadOnlyList<PdfRedactionMatch> matches,
        IReadOnlyList<PdfDiagnosticFinding> findings,
        IReadOnlyList<string>? searchCriteria = null) {
        Preflight = preflight;
        Areas = areas;
        Matches = matches;
        Findings = findings;
        SearchCriteria = searchCriteria ?? Array.Empty<string>();
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

    /// <summary>True when the plan areas were derived from explicit search criteria.</summary>
    public bool IsSearchDriven => SearchCriteria.Count > 0;

    /// <summary>True when at least one match was found.</summary>
    public bool HasMatches => Matches.Count > 0;
}
