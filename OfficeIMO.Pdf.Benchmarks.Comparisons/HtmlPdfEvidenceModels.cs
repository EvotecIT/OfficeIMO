namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal sealed record HtmlPdfEvidenceReport(
    int SchemaVersion,
    DateTimeOffset GeneratedUtc,
    string Scale,
    int Iterations,
    HtmlPdfEvidenceEnvironment Environment,
    HtmlPdfEvidenceProvenance Provenance,
    HtmlPdfEvidenceInput Input,
    IReadOnlyList<HtmlPdfEngineEvidence> Engines);

internal sealed record HtmlPdfEvidenceProvenance(
    HtmlPdfSourceReference OfficeIMO,
    HtmlPdfSourceReference HtmlTinkerX);

internal sealed record HtmlPdfSourceReference(
    string Kind,
    string Version,
    string? Commit,
    bool? WorktreeClean);

internal sealed record HtmlPdfEvidenceEnvironment(
    string OsDescription,
    string OsArchitecture,
    string ProcessArchitecture,
    string RuntimeVersion,
    string FrameworkDescription,
    string? ExternalRasterizer);

internal sealed record HtmlPdfEvidenceInput(
    string RelativePath,
    long SizeBytes,
    string Sha256,
    int ExpectedPageCount,
    int ExpectedReportMarkerCount);

internal sealed record HtmlPdfEngineEvidence(
    string Engine,
    string Owner,
    string AssemblyVersion,
    string ExecutionKind,
    HtmlPdfCancellationEvidence Cancellation,
    HtmlPdfDeterminismEvidence Determinism,
    string MemoryScope,
    bool MemoryComparable,
    IReadOnlyList<HtmlPdfOutputEvidence> Outputs);

internal sealed record HtmlPdfCancellationEvidence(
    bool ApiSupportsCancellation,
    string Status,
    string Detail);

internal sealed record HtmlPdfDeterminismEvidence(
    bool ExactBytesIdentical,
    bool SemanticOutputIdentical,
    bool ManagedVisualPreviewIdentical,
    bool? ExternalVisualPreviewIdentical,
    int UniqueByteHashCount,
    int UniqueSemanticHashCount,
    int UniqueManagedVisualHashCount,
    int? UniqueExternalVisualHashCount);

internal sealed record HtmlPdfOutputEvidence(
    int Iteration,
    string RelativePath,
    double DurationMilliseconds,
    long SizeBytes,
    string Sha256,
    string SemanticSha256,
    long ManagedAllocatedBytes,
    HtmlPdfProcessTreeMemoryEvidence ProcessTreeMemory,
    HtmlPdfContractEvidence Contract,
    HtmlPdfVisualEvidence ManagedVisual,
    HtmlPdfVisualEvidence? ExternalVisual);

internal sealed record HtmlPdfProcessTreeMemoryEvidence(
    long PeakWorkingSetBytes,
    int SampleCount,
    int MinimumObservedProcessCount,
    int MaximumObservedProcessCount,
    string Sampler);

internal sealed record HtmlPdfVisualEvidence(
    string Renderer,
    string RelativePath,
    int PageNumber,
    int Width,
    int Height,
    long SizeBytes,
    string Sha256,
    IReadOnlyList<string> Diagnostics);

internal sealed record HtmlPdfContractEvidence(
    int PageCount,
    int TextLength,
    int ReportMarkerCount,
    long CharacterChecksum,
    bool Tagged,
    bool Marked,
    string? CatalogLanguage,
    int LanguageElementCount,
    int StructureElementCount,
    int MarkedContentReferenceCount,
    int ParentTreeEntryCount,
    bool HasDocumentStructureElement,
    bool FiguresHaveAlternateText,
    IReadOnlyDictionary<string, int> StructureTypeCounts);
