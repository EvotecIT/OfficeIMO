namespace OfficeIMO.Reader.Pdf;

/// <summary>
/// Describes the product contract for the OfficeIMO.Reader PDF adapter.
/// </summary>
public sealed class ReaderPdfProfileContract {
    internal ReaderPdfProfileContract(string id, string displayName, string pipeline, string outputContract, string safetyContract, string unsupportedScope) {
        Id = id;
        DisplayName = displayName;
        Pipeline = pipeline;
        OutputContract = outputContract;
        SafetyContract = safetyContract;
        UnsupportedScope = unsupportedScope;
    }

    /// <summary>Stable profile identifier for manifests, wrappers, and documentation.</summary>
    public string Id { get; }

    /// <summary>Human-readable profile name.</summary>
    public string DisplayName { get; }

    /// <summary>First-party ingestion pipeline used by this profile.</summary>
    public string Pipeline { get; }

    /// <summary>Reader chunk fields and metadata callers can rely on.</summary>
    public string OutputContract { get; }

    /// <summary>Safety and resource behavior callers can rely on.</summary>
    public string SafetyContract { get; }

    /// <summary>Known unsupported or intentionally simplified scope.</summary>
    public string UnsupportedScope { get; }
}

/// <summary>
/// Stable contract description for the OfficeIMO.Reader PDF adapter.
/// </summary>
public static class ReaderPdfProfileContracts {
    private static readonly ReaderPdfProfileContract OfficeIMOContract = new ReaderPdfProfileContract(
        DocumentReaderPdfRegistrationExtensions.HandlerId,
        "OfficeIMO Reader PDF",
        "PDF -> OfficeIMO.Pdf logical model -> DocumentReader page-aware chunks",
        "Emits PDF input kind, normalized source identity, page-aware locations, Markdown text, detected table payloads, image placeholders, link annotations, form widget summaries, chunk hashes, and MaxChars split warnings where requested by ReaderOptions.",
        "Uses Reader input limits, stream-position-aware reads, page-range validation, safe markdown rendering for unsafe links, and parser diagnostics from OfficeIMO.Pdf.",
        "Best for born-digital parser-supported PDFs; scanned PDFs require OCR, and complex encrypted/signed/tagged/active-content PDFs remain subject to OfficeIMO.Pdf read and rewrite blockers.");

    /// <summary>Gets the current OfficeIMO.Reader PDF profile contract.</summary>
    public static ReaderPdfProfileContract OfficeIMO => OfficeIMOContract;
}
