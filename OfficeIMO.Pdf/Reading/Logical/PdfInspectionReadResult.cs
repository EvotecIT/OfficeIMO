namespace OfficeIMO.Pdf;

/// <summary>Outcome of an explicitly requested page inspection.</summary>
public enum PdfPageReadStatus {
    /// <summary>The canonical pipeline produced a logical page.</summary>
    Read,
    /// <summary>Shown composite text had no usable Unicode mapping.</summary>
    UnmappableText,
    /// <summary>A page or invoked Form XObject declared an unsupported content-stream filter.</summary>
    UnsupportedContentStreamFilter
}

/// <summary>One selected page occurrence. A failed page is never represented as successful empty text.</summary>
public sealed class PdfPageReadOutcome {
    internal PdfPageReadOutcome(int pageNumber, int selectionIndex, PdfPageReadStatus status,
        PdfLogicalPage? page = null, PdfUnsupportedTextMappingException? mappingFailure = null,
        IReadOnlyList<string>? unsupportedFilters = null) {
        PageNumber = pageNumber;
        SelectionIndex = selectionIndex;
        Status = status;
        Page = page;
        MappingFailure = mappingFailure;
        UnsupportedFilters = unsupportedFilters ?? Array.Empty<string>();
    }
    /// <summary>One-based source page number.</summary>
    public int PageNumber { get; }
    /// <summary>Zero-based occurrence in the caller's page selection, including repeated pages.</summary>
    public int SelectionIndex { get; }
    /// <summary>Typed completion or failure status.</summary>
    public PdfPageReadStatus Status { get; }
    /// <summary>Canonical logical page, available only after a successful read.</summary>
    public PdfLogicalPage? Page { get; }
    /// <summary>Typed unmappable-font context, when that prevented page extraction.</summary>
    public PdfUnsupportedTextMappingException? MappingFailure { get; }
    /// <summary>Unsupported declared filters for a skipped page.</summary>
    public IReadOnlyList<string> UnsupportedFilters { get; }
    /// <summary>True when this selected page was read successfully.</summary>
    public bool Succeeded => Status == PdfPageReadStatus.Read;
}

/// <summary>Canonical successful page content together with retained document structure and explicit page failures.</summary>
public sealed class PdfInspectionReadResult {
    internal PdfInspectionReadResult(PdfDocumentInfo sourceInfo, PdfRepairReport repairReport,
        PdfDocumentReadResult document, IReadOnlyList<PdfPageReadOutcome> outcomes) {
        SourceInfo = sourceInfo;
        RepairReport = repairReport;
        Document = document;
        PageOutcomes = outcomes;
    }
    /// <summary>Full source structural inventory, including facts on pages whose text was unreadable.</summary>
    public PdfDocumentInfo SourceInfo { get; }
    /// <summary>Structural recoveries and detected-only issues from the canonical parser.</summary>
    public PdfRepairReport RepairReport { get; }
    /// <summary>Canonical logical document containing only successful selected page occurrences.</summary>
    public PdfDocumentReadResult Document { get; }
    /// <summary>Every selected occurrence in caller order.</summary>
    public IReadOnlyList<PdfPageReadOutcome> PageOutcomes { get; }
    /// <summary>True when every selected page completed; this is not a general fidelity or conformance claim.</summary>
    public bool CompletedAllSelectedPages => PageOutcomes.All(static outcome => outcome.Succeeded);
}
