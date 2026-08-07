using System;
using System.Collections.Generic;

namespace OfficeIMO;

/// <summary>Identifies the source format represented by an <see cref="OfficeDocumentModel"/>.</summary>
public enum OfficeDocumentFormat {
    /// <summary>The source format is unknown.</summary>
    Unknown = 0,
    /// <summary>Word document content.</summary>
    Word,
    /// <summary>Excel workbook content.</summary>
    Excel,
    /// <summary>PowerPoint presentation content.</summary>
    PowerPoint,
    /// <summary>Markdown content.</summary>
    Markdown,
    /// <summary>Plain text content.</summary>
    Text,
    /// <summary>PDF content.</summary>
    Pdf,
    /// <summary>Delimited tabular content.</summary>
    Csv,
    /// <summary>JSON content.</summary>
    Json,
    /// <summary>XML content.</summary>
    Xml,
    /// <summary>HTML content.</summary>
    Html,
    /// <summary>Archive content.</summary>
    Zip,
    /// <summary>EPUB content.</summary>
    Epub,
    /// <summary>Visio drawing content.</summary>
    Visio,
    /// <summary>YAML content.</summary>
    Yaml,
    /// <summary>Rich Text Format content.</summary>
    Rtf,
    /// <summary>OpenDocument content.</summary>
    OpenDocument,
    /// <summary>AsciiDoc content.</summary>
    AsciiDoc,
    /// <summary>LaTeX content.</summary>
    Latex,
    /// <summary>Email or MIME content.</summary>
    Email,
    /// <summary>OneNote content.</summary>
    OneNote,
    /// <summary>Calendar content.</summary>
    Calendar,
    /// <summary>Contact-card content.</summary>
    VCard
}

/// <summary>
/// Dependency-free, source-neutral document model used between format engines and conversion destinations.
/// Reader chunks, routing, processing, and transport serialization intentionally remain outside this contract.
/// </summary>
public sealed class OfficeDocumentModel {
    /// <summary>Source format represented by this model.</summary>
    public OfficeDocumentFormat Format { get; set; }

    /// <summary>Source identity and descriptive metadata.</summary>
    public OfficeDocumentModelSource Source { get; set; } = new OfficeDocumentModelSource();

    /// <summary>Capability identifiers used while producing this model.</summary>
    public IReadOnlyList<string> CapabilitiesUsed { get; set; } = Array.Empty<string>();

    /// <summary>Portable Markdown representation when one is available.</summary>
    public string? Markdown { get; set; }

    /// <summary>Portable HTML representation when one is available.</summary>
    public string? Html { get; set; }

    /// <summary>Document-level metadata.</summary>
    public IReadOnlyList<OfficeDocumentModelMetadataEntry> Metadata { get; set; } = Array.Empty<OfficeDocumentModelMetadataEntry>();

    /// <summary>Page-like source containers.</summary>
    public IReadOnlyList<OfficeDocumentModelPage> Pages { get; set; } = Array.Empty<OfficeDocumentModelPage>();

    /// <summary>Logical source blocks in reading order.</summary>
    public IReadOnlyList<OfficeDocumentModelBlock> Blocks { get; set; } = Array.Empty<OfficeDocumentModelBlock>();

    /// <summary>Structured tables.</summary>
    public IReadOnlyList<OfficeDocumentModelTable> Tables { get; set; } = Array.Empty<OfficeDocumentModelTable>();

    /// <summary>Binary or visual assets.</summary>
    public IReadOnlyList<OfficeDocumentModelAsset> Assets { get; set; } = Array.Empty<OfficeDocumentModelAsset>();

    /// <summary>Links and navigation targets.</summary>
    public IReadOnlyList<OfficeDocumentModelLink> Links { get; set; } = Array.Empty<OfficeDocumentModelLink>();

    /// <summary>Form fields and widgets.</summary>
    public IReadOnlyList<OfficeDocumentModelFormField> Forms { get; set; } = Array.Empty<OfficeDocumentModelFormField>();

    /// <summary>Structured visual payloads.</summary>
    public IReadOnlyList<OfficeDocumentModelVisual> Visuals { get; set; } = Array.Empty<OfficeDocumentModelVisual>();

    /// <summary>Loss, fallback, or source diagnostics.</summary>
    public IReadOnlyList<OfficeDocumentModelDiagnostic> Diagnostics { get; set; } = Array.Empty<OfficeDocumentModelDiagnostic>();
}

/// <summary>Source identity and descriptive metadata.</summary>
public sealed class OfficeDocumentModelSource {
    /// <summary>Source path or logical name.</summary>
    public string? Path { get; set; }
    /// <summary>Stable source identifier.</summary>
    public string? SourceId { get; set; }
    /// <summary>Source content hash.</summary>
    public string? SourceHash { get; set; }
    /// <summary>Source last-write timestamp in UTC.</summary>
    public DateTime? LastWriteUtc { get; set; }
    /// <summary>Source length in bytes.</summary>
    public long? LengthBytes { get; set; }
    /// <summary>Document title.</summary>
    public string? Title { get; set; }
    /// <summary>Document author.</summary>
    public string? Author { get; set; }
    /// <summary>Document subject.</summary>
    public string? Subject { get; set; }
    /// <summary>Document keywords.</summary>
    public string? Keywords { get; set; }
}

/// <summary>Source location shared by neutral document elements.</summary>
public sealed class OfficeDocumentModelLocation {
    /// <summary>Source path or logical name.</summary>
    public string? Path { get; set; }
    /// <summary>Zero-based output block index.</summary>
    public int? BlockIndex { get; set; }
    /// <summary>Producer-defined source block index.</summary>
    public int? SourceBlockIndex { get; set; }
    /// <summary>One-based source start line.</summary>
    public int? StartLine { get; set; }
    /// <summary>One-based source end line.</summary>
    public int? EndLine { get; set; }
    /// <summary>One-based normalized start line.</summary>
    public int? NormalizedStartLine { get; set; }
    /// <summary>One-based normalized end line.</summary>
    public int? NormalizedEndLine { get; set; }
    /// <summary>Heading hierarchy path.</summary>
    public string? HeadingPath { get; set; }
    /// <summary>Heading slug or anchor.</summary>
    public string? HeadingSlug { get; set; }
    /// <summary>Producer-defined source block kind.</summary>
    public string? SourceBlockKind { get; set; }
    /// <summary>Deterministic logical block anchor.</summary>
    public string? BlockAnchor { get; set; }
    /// <summary>Spreadsheet sheet name.</summary>
    public string? Sheet { get; set; }
    /// <summary>Spreadsheet A1 range.</summary>
    public string? A1Range { get; set; }
    /// <summary>One-based slide number.</summary>
    public int? Slide { get; set; }
    /// <summary>One-based page number.</summary>
    public int? Page { get; set; }
    /// <summary>Zero-based table index in the closest source container.</summary>
    public int? TableIndex { get; set; }
}

/// <summary>Rectangular source region.</summary>
public sealed class OfficeDocumentModelRegion {
    /// <summary>Left X coordinate.</summary>
    public double X { get; set; }
    /// <summary>Bottom or top Y coordinate according to the producer coordinate system.</summary>
    public double Y { get; set; }
    /// <summary>Region width.</summary>
    public double Width { get; set; }
    /// <summary>Region height.</summary>
    public double Height { get; set; }
}

/// <summary>Page, slide, sheet, or diagram-page container.</summary>
public sealed class OfficeDocumentModelPage {
    /// <summary>One-based source number.</summary>
    public int? Number { get; set; }
    /// <summary>Source name or label.</summary>
    public string? Name { get; set; }
    /// <summary>Plain-text page representation when one is available.</summary>
    public string? Text { get; set; }
    /// <summary>Markdown page representation when one is available.</summary>
    public string? Markdown { get; set; }
    /// <summary>Width in points when known.</summary>
    public double? Width { get; set; }
    /// <summary>Height in points when known.</summary>
    public double? Height { get; set; }
    /// <summary>Rotation in degrees.</summary>
    public int? RotationDegrees { get; set; }
    /// <summary>Source location.</summary>
    public OfficeDocumentModelLocation Location { get; set; } = new OfficeDocumentModelLocation();
    /// <summary>Logical blocks.</summary>
    public IReadOnlyList<OfficeDocumentModelBlock> Blocks { get; set; } = Array.Empty<OfficeDocumentModelBlock>();
    /// <summary>Structured tables.</summary>
    public IReadOnlyList<OfficeDocumentModelTable> Tables { get; set; } = Array.Empty<OfficeDocumentModelTable>();
    /// <summary>Assets.</summary>
    public IReadOnlyList<OfficeDocumentModelAsset> Assets { get; set; } = Array.Empty<OfficeDocumentModelAsset>();
    /// <summary>Links.</summary>
    public IReadOnlyList<OfficeDocumentModelLink> Links { get; set; } = Array.Empty<OfficeDocumentModelLink>();
    /// <summary>Forms.</summary>
    public IReadOnlyList<OfficeDocumentModelFormField> Forms { get; set; } = Array.Empty<OfficeDocumentModelFormField>();
}

/// <summary>Logical source block.</summary>
public sealed class OfficeDocumentModelBlock {
    /// <summary>Stable block identifier.</summary>
    public string Id { get; set; } = string.Empty;
    /// <summary>Producer-normalized block kind.</summary>
    public string Kind { get; set; } = string.Empty;
    /// <summary>Text content.</summary>
    public string Text { get; set; } = string.Empty;
    /// <summary>Heading or list level.</summary>
    public int? Level { get; set; }
    /// <summary>List or leader marker.</summary>
    public string? Marker { get; set; }
    /// <summary>Source location.</summary>
    public OfficeDocumentModelLocation Location { get; set; } = new OfficeDocumentModelLocation();
    /// <summary>Source geometry.</summary>
    public OfficeDocumentModelRegion? Region { get; set; }
}

/// <summary>Structured source table.</summary>
public sealed class OfficeDocumentModelTable {
    /// <summary>Optional table title.</summary>
    public string? Title { get; set; }
    /// <summary>Producer-defined table kind.</summary>
    public string? Kind { get; set; }
    /// <summary>Optional summary.</summary>
    public string? Summary { get; set; }
    /// <summary>Stable payload hash.</summary>
    public string? PayloadHash { get; set; }
    /// <summary>Source location.</summary>
    public OfficeDocumentModelLocation? Location { get; set; }
    /// <summary>Column names.</summary>
    public IReadOnlyList<string> Columns { get; set; } = Array.Empty<string>();
    /// <summary>Rows aligned with <see cref="Columns"/>.</summary>
    public IReadOnlyList<IReadOnlyList<string>> Rows { get; set; } = Array.Empty<IReadOnlyList<string>>();
    /// <summary>Row count before truncation.</summary>
    public int TotalRowCount { get; set; }
    /// <summary>Whether rows were truncated.</summary>
    public bool Truncated { get; set; }
}

/// <summary>Binary or visual asset.</summary>
public sealed class OfficeDocumentModelAsset {
    /// <summary>Stable asset identifier.</summary>
    public string Id { get; set; } = string.Empty;
    /// <summary>Producer-normalized asset kind.</summary>
    public string Kind { get; set; } = string.Empty;
    /// <summary>Media type.</summary>
    public string? MediaType { get; set; }
    /// <summary>Suggested extension.</summary>
    public string? Extension { get; set; }
    /// <summary>Suggested filename.</summary>
    public string? FileName { get; set; }
    /// <summary>Alternative text.</summary>
    public string? AltText { get; set; }
    /// <summary>Title or caption.</summary>
    public string? Title { get; set; }
    /// <summary>Intrinsic width.</summary>
    public int? Width { get; set; }
    /// <summary>Intrinsic height.</summary>
    public int? Height { get; set; }
    /// <summary>Payload length.</summary>
    public long? LengthBytes { get; set; }
    /// <summary>Stable payload hash.</summary>
    public string? PayloadHash { get; set; }
    /// <summary>Optional in-memory payload.</summary>
    public byte[]? PayloadBytes { get; set; }
    /// <summary>Source object identifier.</summary>
    public string? SourceObjectId { get; set; }
    /// <summary>Source geometry.</summary>
    public OfficeDocumentModelRegion? Region { get; set; }
    /// <summary>Source location.</summary>
    public OfficeDocumentModelLocation Location { get; set; } = new OfficeDocumentModelLocation();
}

/// <summary>Hyperlink or navigation target.</summary>
public sealed class OfficeDocumentModelLink {
    /// <summary>Stable link identifier.</summary>
    public string Id { get; set; } = string.Empty;
    /// <summary>Link kind.</summary>
    public string Kind { get; set; } = string.Empty;
    /// <summary>URI target.</summary>
    public string? Uri { get; set; }
    /// <summary>Internal destination name.</summary>
    public string? DestinationName { get; set; }
    /// <summary>Internal destination page.</summary>
    public int? DestinationPageNumber { get; set; }
    /// <summary>Destination mode.</summary>
    public string? DestinationMode { get; set; }
    /// <summary>Viewer named action.</summary>
    public string? NamedAction { get; set; }
    /// <summary>Remote file target.</summary>
    public string? RemoteFile { get; set; }
    /// <summary>Remote destination name.</summary>
    public string? RemoteDestinationName { get; set; }
    /// <summary>Remote destination page.</summary>
    public int? RemoteDestinationPageNumber { get; set; }
    /// <summary>Display text.</summary>
    public string? Text { get; set; }
    /// <summary>Source location.</summary>
    public OfficeDocumentModelLocation Location { get; set; } = new OfficeDocumentModelLocation();
    /// <summary>Source geometry.</summary>
    public OfficeDocumentModelRegion? Region { get; set; }
}

/// <summary>Form field or widget.</summary>
public sealed class OfficeDocumentModelFormField {
    /// <summary>Stable field identifier.</summary>
    public string Id { get; set; } = string.Empty;
    /// <summary>Field name.</summary>
    public string? Name { get; set; }
    /// <summary>Field kind.</summary>
    public string Kind { get; set; } = string.Empty;
    /// <summary>Current value.</summary>
    public string? Value { get; set; }
    /// <summary>Whether the field is read-only.</summary>
    public bool IsReadOnly { get; set; }
    /// <summary>Whether the field is required.</summary>
    public bool IsRequired { get; set; }
    /// <summary>Source location.</summary>
    public OfficeDocumentModelLocation Location { get; set; } = new OfficeDocumentModelLocation();
    /// <summary>Source geometry.</summary>
    public OfficeDocumentModelRegion? Region { get; set; }
}

/// <summary>Structured visual payload.</summary>
public sealed class OfficeDocumentModelVisual {
    /// <summary>Normalized visual kind.</summary>
    public string Kind { get; set; } = string.Empty;
    /// <summary>Payload language.</summary>
    public string Language { get; set; } = string.Empty;
    /// <summary>Raw visual payload.</summary>
    public string Content { get; set; } = string.Empty;
    /// <summary>Stable payload hash.</summary>
    public string? PayloadHash { get; set; }
    /// <summary>Source name.</summary>
    public string? SourceName { get; set; }
    /// <summary>Media type.</summary>
    public string? MediaType { get; set; }
    /// <summary>Intrinsic width.</summary>
    public double? Width { get; set; }
    /// <summary>Intrinsic height.</summary>
    public double? Height { get; set; }
    /// <summary>Placed X coordinate.</summary>
    public double? X { get; set; }
    /// <summary>Placed Y coordinate.</summary>
    public double? Y { get; set; }
    /// <summary>Placed width.</summary>
    public double? PlacedWidth { get; set; }
    /// <summary>Placed height.</summary>
    public double? PlacedHeight { get; set; }
    /// <summary>Placement count.</summary>
    public int PlacementCount { get; set; }
    /// <summary>Whether placement geometry is available.</summary>
    public bool HasGeometry { get; set; }
    /// <summary>Whether placement geometry is axis-aligned.</summary>
    public bool? IsAxisAligned { get; set; }
    /// <summary>Source location.</summary>
    public OfficeDocumentModelLocation? Location { get; set; }
}

/// <summary>Document metadata entry.</summary>
public sealed class OfficeDocumentModelMetadataEntry {
    /// <summary>Stable metadata identifier.</summary>
    public string Id { get; set; } = string.Empty;
    /// <summary>Metadata category.</summary>
    public string Category { get; set; } = string.Empty;
    /// <summary>Metadata name.</summary>
    public string Name { get; set; } = string.Empty;
    /// <summary>Metadata value.</summary>
    public string? Value { get; set; }
    /// <summary>Value kind.</summary>
    public string? ValueType { get; set; }
    /// <summary>Source object identifier.</summary>
    public string? SourceObjectId { get; set; }
    /// <summary>Source location.</summary>
    public OfficeDocumentModelLocation? Location { get; set; }
    /// <summary>Additional scalar attributes.</summary>
    public IReadOnlyDictionary<string, string> Attributes { get; set; } = new Dictionary<string, string>(StringComparer.Ordinal);
}

/// <summary>Diagnostic emitted while constructing or consuming a neutral model.</summary>
public sealed class OfficeDocumentModelDiagnostic {
    /// <summary>Diagnostic severity.</summary>
    public OfficeDocumentModelDiagnosticSeverity Severity { get; set; } = OfficeDocumentModelDiagnosticSeverity.Warning;
    /// <summary>Diagnostic category.</summary>
    public OfficeDocumentModelDiagnosticCategory Category { get; set; }
    /// <summary>Stable diagnostic code.</summary>
    public string Code { get; set; } = string.Empty;
    /// <summary>Human-readable message.</summary>
    public string Message { get; set; } = string.Empty;
    /// <summary>Producing component.</summary>
    public string? Source { get; set; }
    /// <summary>Whether processing can continue.</summary>
    public bool? IsRecoverable { get; set; }
    /// <summary>Source location.</summary>
    public OfficeDocumentModelLocation? Location { get; set; }
    /// <summary>Machine-readable details.</summary>
    public IReadOnlyDictionary<string, string> Attributes { get; set; } = new Dictionary<string, string>(StringComparer.Ordinal);
}

/// <summary>Neutral diagnostic category.</summary>
public enum OfficeDocumentModelDiagnosticCategory {
    /// <summary>Unclassified diagnostic.</summary>
    General = 0,
    /// <summary>Input detection.</summary>
    Detection,
    /// <summary>Input access or integrity.</summary>
    Input,
    /// <summary>Format parsing.</summary>
    Parsing,
    /// <summary>Content loss or fallback.</summary>
    Content,
    /// <summary>Security behavior.</summary>
    Security,
    /// <summary>Configured limit.</summary>
    Limit,
    /// <summary>Format-specific behavior.</summary>
    Adapter
}

/// <summary>Neutral diagnostic severity.</summary>
public enum OfficeDocumentModelDiagnosticSeverity {
    /// <summary>Informational diagnostic.</summary>
    Information,
    /// <summary>Recoverable warning.</summary>
    Warning,
    /// <summary>Error or incomplete conversion.</summary>
    Error
}
