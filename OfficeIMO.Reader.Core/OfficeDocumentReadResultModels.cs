using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace OfficeIMO.Reader;

/// <summary>
/// Stable schema values for the OfficeIMO document read result transport contract.
/// </summary>
public static partial class OfficeDocumentReadResultSchema {
    /// <summary>
    /// Stable schema identifier for shared document readback results.
    /// </summary>
    public const string Id = "officeimo.document.read-result";

}

/// <summary>
/// Shared result envelope for document readback workflows.
/// </summary>
public sealed class OfficeDocumentReadResult {
    /// <summary>
    /// Result schema identifier.
    /// </summary>
    public string SchemaId { get; set; } = OfficeDocumentReadResultSchema.Id;

    /// <summary>
    /// Result schema version.
    /// </summary>
    public int SchemaVersion { get; set; } = OfficeDocumentReadResultSchema.CurrentVersion;

    /// <summary>
    /// Source input kind that produced this result.
    /// </summary>
    public ReaderInputKind Kind { get; set; } = ReaderInputKind.Unknown;

    /// <summary>
    /// Source document metadata used for citations and incremental processing.
    /// </summary>
    public OfficeDocumentSource Source { get; set; } = new OfficeDocumentSource();

    /// <summary>
    /// Capability identifiers used while producing this result.
    /// </summary>
    public IReadOnlyList<string> CapabilitiesUsed { get; set; } = Array.Empty<string>();

    /// <summary>
    /// Portable Markdown representation when available.
    /// </summary>
    public string? Markdown { get; set; }

    /// <summary>
    /// Portable or review HTML representation when available.
    /// </summary>
    public string? Html { get; set; }

    /// <summary>
    /// Deterministic JSON representation when a producer emits one.
    /// </summary>
    public string? Json { get; set; }

    /// <summary>
    /// Reader chunks produced by the same extraction run.
    /// </summary>
    public IReadOnlyList<ReaderChunk> Chunks { get; set; } = Array.Empty<ReaderChunk>();

    /// <summary>
    /// Document-level metadata entries discovered during reading.
    /// </summary>
    public IReadOnlyList<OfficeDocumentMetadataEntry> Metadata { get; set; } = Array.Empty<OfficeDocumentMetadataEntry>();

    /// <summary>
    /// Page, slide, sheet, or diagram-page records produced by document reading.
    /// </summary>
    public IReadOnlyList<OfficeDocumentPage> Pages { get; set; } = Array.Empty<OfficeDocumentPage>();

    /// <summary>
    /// Normalized logical blocks in source order.
    /// </summary>
    public IReadOnlyList<OfficeDocumentBlock> Blocks { get; set; } = Array.Empty<OfficeDocumentBlock>();

    /// <summary>
    /// Structured tables extracted or detected during reading.
    /// </summary>
    public IReadOnlyList<ReaderTable> Tables { get; set; } = Array.Empty<ReaderTable>();

    /// <summary>
    /// Images, previews, and other binary or visual assets discovered during reading.
    /// </summary>
    public IReadOnlyList<OfficeDocumentAsset> Assets { get; set; } = Array.Empty<OfficeDocumentAsset>();

    /// <summary>
    /// Hyperlinks, internal destinations, and navigation targets discovered during reading.
    /// </summary>
    public IReadOnlyList<OfficeDocumentLink> Links { get; set; } = Array.Empty<OfficeDocumentLink>();

    /// <summary>
    /// Form fields or widgets discovered during reading.
    /// </summary>
    public IReadOnlyList<OfficeDocumentFormField> Forms { get; set; } = Array.Empty<OfficeDocumentFormField>();

    /// <summary>
    /// Regions that likely need OCR before text extraction can be considered complete.
    /// </summary>
    public IReadOnlyList<OfficeDocumentOcrCandidate> OcrCandidates { get; set; } = Array.Empty<OfficeDocumentOcrCandidate>();

    /// <summary>
    /// Structured visual payloads discovered during reading.
    /// </summary>
    public IReadOnlyList<ReaderVisual> Visuals { get; set; } = Array.Empty<ReaderVisual>();

    /// <summary>
    /// Warnings and diagnostics emitted during reading.
    /// </summary>
    public IReadOnlyList<OfficeDocumentDiagnostic> Diagnostics { get; set; } = Array.Empty<OfficeDocumentDiagnostic>();
}

/// <summary>
/// Normalized document-level metadata entry emitted by a document read adapter.
/// </summary>
public sealed class OfficeDocumentMetadataEntry {
    /// <summary>
    /// Stable metadata identifier within the read result.
    /// </summary>
    public string Id { get; set; } = string.Empty;

    /// <summary>
    /// Producer-normalized metadata category, such as core, catalog, outline, or destination.
    /// </summary>
    public string Category { get; set; } = string.Empty;

    /// <summary>
    /// Metadata entry name.
    /// </summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>
    /// Metadata entry value when it can be represented as stable text.
    /// </summary>
    public string? Value { get; set; }

    /// <summary>
    /// Producer-normalized value kind, such as string, number, boolean, count, or object.
    /// </summary>
    public string? ValueType { get; set; }

    /// <summary>
    /// Source-specific object, relationship, or resource identifier.
    /// </summary>
    public string? SourceObjectId { get; set; }

    /// <summary>
    /// Source location for metadata that targets a document region or page.
    /// </summary>
    public ReaderLocation? Location { get; set; }

    /// <summary>
    /// Additional stable scalar attributes for metadata that needs more than one value.
    /// </summary>
    public IReadOnlyDictionary<string, string> Attributes { get; set; } = new Dictionary<string, string>(StringComparer.Ordinal);
}

/// <summary>
/// Source metadata shared by document read results.
/// </summary>
public sealed class OfficeDocumentSource {
    /// <summary>
    /// Source path or logical source name.
    /// </summary>
    public string? Path { get; set; }

    /// <summary>
    /// Stable source identifier when available.
    /// </summary>
    public string? SourceId { get; set; }

    /// <summary>
    /// Source content hash when available.
    /// </summary>
    public string? SourceHash { get; set; }

    /// <summary>
    /// Source last-write timestamp in UTC when available.
    /// </summary>
    public DateTime? LastWriteUtc { get; set; }

    /// <summary>
    /// Source length in bytes when available.
    /// </summary>
    public long? LengthBytes { get; set; }

    /// <summary>
    /// Source title metadata when available.
    /// </summary>
    public string? Title { get; set; }

    /// <summary>
    /// Source author metadata when available.
    /// </summary>
    public string? Author { get; set; }

    /// <summary>
    /// Source subject metadata when available.
    /// </summary>
    public string? Subject { get; set; }

    /// <summary>
    /// Source keyword metadata when available.
    /// </summary>
    public string? Keywords { get; set; }
}

/// <summary>
/// Page-like container in a shared read result.
/// </summary>
public sealed class OfficeDocumentPage {
    /// <summary>
    /// One-based page, slide, sheet, or diagram-page number when applicable.
    /// </summary>
    public int? Number { get; set; }

    /// <summary>
    /// Producer-specific page name or label when available.
    /// </summary>
    public string? Name { get; set; }

    /// <summary>
    /// Page width in points when available.
    /// </summary>
    public double? Width { get; set; }

    /// <summary>
    /// Page height in points when available.
    /// </summary>
    public double? Height { get; set; }

    /// <summary>
    /// Rotation in degrees when available.
    /// </summary>
    public int? RotationDegrees { get; set; }

    /// <summary>
    /// Location metadata for the page-like container.
    /// </summary>
    public ReaderLocation Location { get; set; } = new ReaderLocation();

    /// <summary>
    /// Logical blocks on this page-like container.
    /// </summary>
    public IReadOnlyList<OfficeDocumentBlock> Blocks { get; set; } = Array.Empty<OfficeDocumentBlock>();

    /// <summary>
    /// Structured tables associated with this page-like container.
    /// </summary>
    public IReadOnlyList<ReaderTable> Tables { get; set; } = Array.Empty<ReaderTable>();

    /// <summary>
    /// Assets associated with this page-like container.
    /// </summary>
    public IReadOnlyList<OfficeDocumentAsset> Assets { get; set; } = Array.Empty<OfficeDocumentAsset>();

    /// <summary>
    /// Links associated with this page-like container.
    /// </summary>
    public IReadOnlyList<OfficeDocumentLink> Links { get; set; } = Array.Empty<OfficeDocumentLink>();

    /// <summary>
    /// Form fields or widgets associated with this page-like container.
    /// </summary>
    public IReadOnlyList<OfficeDocumentFormField> Forms { get; set; } = Array.Empty<OfficeDocumentFormField>();

    /// <summary>
    /// OCR candidate regions associated with this page-like container.
    /// </summary>
    public IReadOnlyList<OfficeDocumentOcrCandidate> OcrCandidates { get; set; } = Array.Empty<OfficeDocumentOcrCandidate>();
}

/// <summary>
/// Normalized logical block emitted by a document read adapter.
/// </summary>
public sealed class OfficeDocumentBlock {
    /// <summary>
    /// Stable block identifier within the read result.
    /// </summary>
    public string Id { get; set; } = string.Empty;

    /// <summary>
    /// Producer-normalized block kind, such as heading, paragraph, table, or list-item.
    /// </summary>
    public string Kind { get; set; } = string.Empty;

    /// <summary>
    /// Text content for text-like blocks.
    /// </summary>
    public string Text { get; set; } = string.Empty;

    /// <summary>
    /// Optional block level, such as heading level or list nesting depth.
    /// </summary>
    public int? Level { get; set; }

    /// <summary>
    /// Optional marker for list or leader blocks.
    /// </summary>
    public string? Marker { get; set; }

    /// <summary>
    /// Source location for this block.
    /// </summary>
    public ReaderLocation Location { get; set; } = new ReaderLocation();

    /// <summary>
    /// Source geometry in points when available.
    /// </summary>
    public OfficeDocumentRegion? Region { get; set; }
}

/// <summary>
/// Rectangular source region in document coordinate units.
/// </summary>
public sealed class OfficeDocumentRegion {
    /// <summary>Left X coordinate.</summary>
    public double X { get; set; }

    /// <summary>Bottom or top Y coordinate, depending on the source coordinate system.</summary>
    public double Y { get; set; }

    /// <summary>Region width.</summary>
    public double Width { get; set; }

    /// <summary>Region height.</summary>
    public double Height { get; set; }
}

/// <summary>
/// Asset metadata emitted by a document read adapter.
/// </summary>
public sealed class OfficeDocumentAsset {
    /// <summary>
    /// Stable asset identifier within the read result.
    /// </summary>
    public string Id { get; set; } = string.Empty;

    /// <summary>
    /// Normalized asset kind, such as image or preview.
    /// </summary>
    public string Kind { get; set; } = string.Empty;

    /// <summary>
    /// Media type when known.
    /// </summary>
    public string? MediaType { get; set; }

    /// <summary>
    /// Suggested file extension when known.
    /// </summary>
    public string? Extension { get; set; }

    /// <summary>
    /// Deterministic suggested filename for writing this asset outside the read result envelope.
    /// </summary>
    public string? FileName { get; set; }

    /// <summary>
    /// Accessibility description or alternate text when available from the source document.
    /// </summary>
    public string? AltText { get; set; }

    /// <summary>
    /// Source title or caption-like label when available from the asset metadata.
    /// </summary>
    public string? Title { get; set; }

    /// <summary>
    /// Width in pixels or source units when known.
    /// </summary>
    public int? Width { get; set; }

    /// <summary>
    /// Height in pixels or source units when known.
    /// </summary>
    public int? Height { get; set; }

    /// <summary>
    /// Asset byte length when known.
    /// </summary>
    public long? LengthBytes { get; set; }

    /// <summary>
    /// Stable payload hash when available.
    /// </summary>
    public string? PayloadHash { get; set; }

    /// <summary>
    /// Optional in-memory payload for callers that request materializable assets. This payload is not included in JSON transport output.
    /// </summary>
    [JsonIgnore]
    public byte[]? PayloadBytes { get; set; }

    /// <summary>
    /// Source-specific relationship, resource, or object identifier.
    /// </summary>
    public string? SourceObjectId { get; set; }

    /// <summary>
    /// Source geometry for this asset when the read adapter can locate a concrete placement.
    /// </summary>
    public OfficeDocumentRegion? Region { get; set; }

    /// <summary>
    /// Source location for this asset.
    /// </summary>
    public ReaderLocation Location { get; set; } = new ReaderLocation();
}

/// <summary>
/// Hyperlink or navigation target emitted by a document read adapter.
/// </summary>
public sealed class OfficeDocumentLink {
    /// <summary>
    /// Stable link identifier within the read result.
    /// </summary>
    public string Id { get; set; } = string.Empty;

    /// <summary>
    /// Normalized link kind, such as uri, destination, named-action, or remote.
    /// </summary>
    public string Kind { get; set; } = string.Empty;

    /// <summary>
    /// URI target when available.
    /// </summary>
    public string? Uri { get; set; }

    /// <summary>
    /// Internal destination name when available.
    /// </summary>
    public string? DestinationName { get; set; }

    /// <summary>
    /// Internal destination page number when available.
    /// </summary>
    public int? DestinationPageNumber { get; set; }

    /// <summary>
    /// Internal destination mode when available.
    /// </summary>
    public string? DestinationMode { get; set; }

    /// <summary>
    /// Internal destination top coordinate when available.
    /// </summary>
    public double? DestinationTop { get; set; }

    /// <summary>
    /// Internal destination left coordinate when available.
    /// </summary>
    public double? DestinationLeft { get; set; }

    /// <summary>
    /// Internal destination bottom coordinate when available.
    /// </summary>
    public double? DestinationBottom { get; set; }

    /// <summary>
    /// Internal destination right coordinate when available.
    /// </summary>
    public double? DestinationRight { get; set; }

    /// <summary>
    /// Viewer named action when available.
    /// </summary>
    public string? NamedAction { get; set; }

    /// <summary>
    /// Remote file target when available.
    /// </summary>
    public string? RemoteFile { get; set; }

    /// <summary>
    /// Remote destination name within the remote file when available.
    /// </summary>
    public string? RemoteDestinationName { get; set; }

    /// <summary>
    /// Remote destination page number within the remote file when available.
    /// </summary>
    public int? RemoteDestinationPageNumber { get; set; }

    /// <summary>
    /// Remote destination mode when available.
    /// </summary>
    public string? RemoteDestinationMode { get; set; }

    /// <summary>
    /// Remote destination top coordinate when available.
    /// </summary>
    public double? RemoteDestinationTop { get; set; }

    /// <summary>
    /// Remote destination left coordinate when available.
    /// </summary>
    public double? RemoteDestinationLeft { get; set; }

    /// <summary>
    /// Remote destination bottom coordinate when available.
    /// </summary>
    public double? RemoteDestinationBottom { get; set; }

    /// <summary>
    /// Remote destination right coordinate when available.
    /// </summary>
    public double? RemoteDestinationRight { get; set; }

    /// <summary>
    /// Optional display or annotation text for the link.
    /// </summary>
    public string? Text { get; set; }

    /// <summary>
    /// Source location for this link.
    /// </summary>
    public ReaderLocation Location { get; set; } = new ReaderLocation();

    /// <summary>
    /// Source geometry in points when available.
    /// </summary>
    public OfficeDocumentRegion? Region { get; set; }
}

/// <summary>
/// Form field or widget emitted by a document read adapter.
/// </summary>
public sealed class OfficeDocumentFormField {
    /// <summary>
    /// Stable form-field identifier within the read result.
    /// </summary>
    public string Id { get; set; } = string.Empty;

    /// <summary>
    /// Field name when available.
    /// </summary>
    public string? Name { get; set; }

    /// <summary>
    /// Normalized or source-specific field kind.
    /// </summary>
    public string Kind { get; set; } = string.Empty;

    /// <summary>
    /// Current field value when available.
    /// </summary>
    public string? Value { get; set; }

    /// <summary>
    /// True when the source marks the field read-only.
    /// </summary>
    public bool IsReadOnly { get; set; }

    /// <summary>
    /// True when the source marks the field required.
    /// </summary>
    public bool IsRequired { get; set; }

    /// <summary>
    /// Source location for this field or widget.
    /// </summary>
    public ReaderLocation Location { get; set; } = new ReaderLocation();

    /// <summary>
    /// Source geometry in points when available.
    /// </summary>
    public OfficeDocumentRegion? Region { get; set; }
}

/// <summary>
/// Region or page that likely needs OCR enrichment.
/// </summary>
public sealed class OfficeDocumentOcrCandidate {
    /// <summary>
    /// Stable OCR candidate identifier within the read result.
    /// </summary>
    public string Id { get; set; } = string.Empty;

    /// <summary>
    /// Candidate kind, such as page or image.
    /// </summary>
    public string Kind { get; set; } = string.Empty;

    /// <summary>
    /// Human-readable reason why OCR should be considered.
    /// </summary>
    public string? Reason { get; set; }

    /// <summary>
    /// Optional confidence that OCR is needed.
    /// </summary>
    public double? Confidence { get; set; }

    /// <summary>
    /// Related asset identifier when the candidate maps to an image asset.
    /// </summary>
    public string? AssetId { get; set; }

    /// <summary>
    /// Number of image resources contributing to this candidate.
    /// </summary>
    public int? ImageCount { get; set; }

    /// <summary>
    /// Number of native text blocks found in the same region or page.
    /// </summary>
    public int? TextBlockCount { get; set; }

    /// <summary>
    /// Source location for this OCR candidate.
    /// </summary>
    public ReaderLocation Location { get; set; } = new ReaderLocation();

    /// <summary>
    /// Source geometry in points when available.
    /// </summary>
    public OfficeDocumentRegion? Region { get; set; }
}

/// <summary>
/// Diagnostic emitted by a document read adapter.
/// </summary>
public sealed class OfficeDocumentDiagnostic {
    /// <summary>
    /// Diagnostic severity.
    /// </summary>
    public OfficeDocumentDiagnosticSeverity Severity { get; set; } = OfficeDocumentDiagnosticSeverity.Warning;

    /// <summary>
    /// Stable diagnostic category for host filtering.
    /// </summary>
    public OfficeDocumentDiagnosticCategory Category { get; set; } = OfficeDocumentDiagnosticCategory.General;

    /// <summary>
    /// Stable diagnostic code.
    /// </summary>
    public string Code { get; set; } = string.Empty;

    /// <summary>
    /// Human-readable diagnostic message.
    /// </summary>
    public string Message { get; set; } = string.Empty;

    /// <summary>
    /// Optional component or adapter that emitted the diagnostic.
    /// </summary>
    public string? Source { get; set; }

    /// <summary>
    /// Whether processing can continue without caller intervention, when known.
    /// </summary>
    public bool? IsRecoverable { get; set; }

    /// <summary>
    /// Optional source location for the diagnostic.
    /// </summary>
    public ReaderLocation? Location { get; set; }

    /// <summary>
    /// Stable machine-readable diagnostic details.
    /// </summary>
    public IReadOnlyDictionary<string, string> Attributes { get; set; } = new Dictionary<string, string>();
}

/// <summary>
/// Stable categories for document diagnostics.
/// </summary>
public enum OfficeDocumentDiagnosticCategory {
    /// <summary>Unclassified diagnostic.</summary>
    General = 0,
    /// <summary>Input kind or content detection.</summary>
    Detection,
    /// <summary>Input access, size, or integrity.</summary>
    Input,
    /// <summary>Parsing or format interpretation.</summary>
    Parsing,
    /// <summary>Content loss, truncation, or unsupported content.</summary>
    Content,
    /// <summary>Security, active content, encryption, or signatures.</summary>
    Security,
    /// <summary>Optical character recognition readiness.</summary>
    Ocr,
    /// <summary>Configured execution or output limit.</summary>
    Limit,
    /// <summary>Format adapter-specific diagnostic.</summary>
    Adapter
}

/// <summary>
/// Severity values for document read diagnostics.
/// </summary>
public enum OfficeDocumentDiagnosticSeverity {
    /// <summary>Informational diagnostic.</summary>
    Information,

    /// <summary>Recoverable warning.</summary>
    Warning,

    /// <summary>Error diagnostic for failed or incomplete document reading.</summary>
    Error
}
