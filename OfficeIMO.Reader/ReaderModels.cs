using OfficeIMO.Markdown;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace OfficeIMO.Reader;

/// <summary>
/// The detected input kind used by <see cref="DocumentReader"/>.
/// </summary>
public enum ReaderInputKind {
    /// <summary>
    /// Unknown/other file kind treated as plain text.
    /// </summary>
    Unknown = 0,
    /// <summary>
    /// Word document (DOCX/DOCM).
    /// </summary>
    Word,
    /// <summary>
    /// Excel workbook (XLSX/XLSM).
    /// </summary>
    Excel,
    /// <summary>
    /// PowerPoint presentation (PPTX/PPTM).
    /// </summary>
    PowerPoint,
    /// <summary>
    /// Markdown text file.
    /// </summary>
    Markdown,
    /// <summary>
    /// Generic text file.
    /// </summary>
    Text,
    /// <summary>
    /// PDF document.
    /// </summary>
    Pdf,
    /// <summary>
    /// CSV/TSV structured text.
    /// </summary>
    Csv,
    /// <summary>
    /// JSON structured text.
    /// </summary>
    Json,
    /// <summary>
    /// XML structured text.
    /// </summary>
    Xml,
    /// <summary>
    /// HTML document.
    /// </summary>
    Html,
    /// <summary>
    /// ZIP archive.
    /// </summary>
    Zip,
    /// <summary>
    /// EPUB e-book archive.
    /// </summary>
    Epub,
    /// <summary>
    /// Visio drawing (VSDX/VSDM/VSTX/VSTM).
    /// </summary>
    Visio,
    /// <summary>
    /// YAML structured text.
    /// </summary>
    Yaml
}

/// <summary>
/// A normalized extraction chunk produced by <see cref="DocumentReader"/>.
/// </summary>
public sealed class ReaderChunk {
    /// <summary>
    /// Stable, ASCII-only identifier.
    /// </summary>
    public string Id { get; set; } = string.Empty;

    /// <summary>
    /// The kind of input that produced this chunk.
    /// </summary>
    public ReaderInputKind Kind { get; set; }

    /// <summary>
    /// Source location information for citations and debugging.
    /// </summary>
    public ReaderLocation Location { get; set; } = new ReaderLocation();

    /// <summary>
    /// Stable identifier for the source document.
    /// For file-based reads this is deterministic for a given normalized path.
    /// </summary>
    public string? SourceId { get; set; }

    /// <summary>
    /// Optional content hash for the source document (for incremental upserts).
    /// </summary>
    public string? SourceHash { get; set; }

    /// <summary>
    /// Optional content hash for this chunk (for incremental upserts).
    /// </summary>
    public string? ChunkHash { get; set; }

    /// <summary>
    /// Optional source last-write timestamp (UTC) when available.
    /// </summary>
    public DateTime? SourceLastWriteUtc { get; set; }

    /// <summary>
    /// Optional source length in bytes when available.
    /// </summary>
    public long? SourceLengthBytes { get; set; }

    /// <summary>
    /// Estimated token count (best-effort heuristic) for prompt budgeting.
    /// </summary>
    public int? TokenEstimate { get; set; }

    /// <summary>
    /// Plain text representation of the chunk.
    /// </summary>
    public string Text { get; set; } = string.Empty;

    /// <summary>
    /// Optional Markdown representation of the chunk.
    /// </summary>
    public string? Markdown { get; set; }

    /// <summary>
    /// Optional structured tables extracted from this chunk.
    /// </summary>
    public IReadOnlyList<ReaderTable>? Tables { get; set; }

    /// <summary>
    /// Optional structured visual fence metadata extracted from this chunk.
    /// </summary>
    public IReadOnlyList<ReaderVisual>? Visuals { get; set; }

    /// <summary>
    /// Optional structured form fields extracted from this chunk.
    /// </summary>
    public IReadOnlyList<ReaderFormField>? FormFields { get; set; }

    /// <summary>
    /// Optional passive action summaries extracted from this chunk.
    /// </summary>
    public IReadOnlyList<ReaderActionSummary>? Actions { get; set; }

    /// <summary>
    /// Optional source-specific diagnostics and structural counters for this chunk.
    /// </summary>
    public ReaderChunkDiagnostics? Diagnostics { get; set; }

    /// <summary>
    /// Optional warnings about truncation or unsupported content.
    /// </summary>
    public IReadOnlyList<string>? Warnings { get; set; }
}

/// <summary>
/// Optional source diagnostics and structural counters attached to a reader chunk.
/// </summary>
public sealed class ReaderChunkDiagnostics {
    /// <summary>Source adapter that produced the diagnostics, for example "pdf".</summary>
    public string SourceKind { get; set; } = string.Empty;

    /// <summary>Total logical page count in the loaded source document.</summary>
    public int PageCount { get; set; }

    /// <summary>Number of pages selected for this read operation.</summary>
    public int SelectedPageCount { get; set; }

    /// <summary>One-based PDF page number for page-scoped chunks, when applicable.</summary>
    public int? PageNumber { get; set; }

    /// <summary>Number of logical tables in the selected scope.</summary>
    public int TableCount { get; set; }

    /// <summary>Number of logical tables in the selected scope that expose placement geometry.</summary>
    public int TableGeometryCount { get; set; }

    /// <summary>Ratio of tables with placement geometry to all tables in the selected scope.</summary>
    public double TableGeometryCoverage { get; set; }

    /// <summary>Lowest table detection confidence in the selected scope, when tables are present.</summary>
    public double? MinTableConfidence { get; set; }

    /// <summary>Average table detection confidence in the selected scope, when tables are present.</summary>
    public double? AverageTableConfidence { get; set; }

    /// <summary>Number of logical images in the selected scope.</summary>
    public int ImageCount { get; set; }

    /// <summary>Number of logical images in the selected scope that expose placement geometry.</summary>
    public int ImageGeometryCount { get; set; }

    /// <summary>Ratio of images with placement geometry to all images in the selected scope.</summary>
    public double ImageGeometryCoverage { get; set; }

    /// <summary>Number of logical link annotations in the selected scope.</summary>
    public int LinkCount { get; set; }

    /// <summary>True when the source exposes a readable document open action.</summary>
    public bool HasOpenAction { get; set; }

    /// <summary>True when the source exposes active catalog-level actions.</summary>
    public bool HasCatalogActions { get; set; }

    /// <summary>True when the selected chunk scope exposes page-level actions.</summary>
    public bool HasPageActions { get; set; }

    /// <summary>True when the selected chunk scope exposes annotation-level actions.</summary>
    public bool HasAnnotationActions { get; set; }

    /// <summary>True when the source exposes active catalog, selected page, or selected annotation actions.</summary>
    public bool HasActiveContent { get; set; }

    /// <summary>Number of active catalog-level actions in the loaded source document.</summary>
    public int CatalogActionCount { get; set; }

    /// <summary>Number of page-level actions in the loaded source document.</summary>
    public int PageActionCount { get; set; }

    /// <summary>Number of page-level actions in this chunk's selected page scope.</summary>
    public int SelectedPageActionCount { get; set; }

    /// <summary>Number of annotation-level actions in the loaded source document.</summary>
    public int AnnotationActionCount { get; set; }

    /// <summary>Number of annotation-level actions in this chunk's selected page scope.</summary>
    public int SelectedAnnotationActionCount { get; set; }

    /// <summary>Total form field count in the loaded source document.</summary>
    public int FormFieldCount { get; set; }

    /// <summary>Total form widget count in the loaded source document.</summary>
    public int FormWidgetCount { get; set; }

    /// <summary>Number of form widgets in this chunk's selected page scope.</summary>
    public int SelectedFormWidgetCount { get; set; }

    /// <summary>Number of selected form widgets that expose a current appearance state.</summary>
    public int SelectedFormWidgetAppearanceStateCount { get; set; }

    /// <summary>Ratio of selected form widgets with a current appearance state to all selected form widgets.</summary>
    public double SelectedFormWidgetAppearanceStateCoverage { get; set; }

    /// <summary>Total readable normal appearance states across selected form widgets.</summary>
    public int SelectedFormWidgetNormalAppearanceStateCount { get; set; }

    /// <summary>True when the source exposes encryption, signature, permission, or incremental-update markers.</summary>
    public bool HasSecurityState { get; set; }

    /// <summary>True when the source PDF contains an encryption marker.</summary>
    public bool HasEncryption { get; set; }

    /// <summary>True when the source PDF contains signature markers, fields, or values.</summary>
    public bool HasSignatures { get; set; }

    /// <summary>True when the source PDF contains incremental-update markers.</summary>
    public bool HasIncrementalUpdates { get; set; }

    /// <summary>Number of readable PDF revision markers.</summary>
    public int RevisionCount { get; set; }

    /// <summary>True when mutation should preserve the existing PDF by appending a new revision.</summary>
    public bool RequiresAppendOnlyMutation { get; set; }
}

/// <summary>
/// Generic location metadata used across supported formats.
/// </summary>
public sealed class ReaderLocation {
    /// <summary>
    /// Source path used for citations.
    /// </summary>
    public string? Path { get; set; }

    /// <summary>
    /// Optional emitted chunk index (0-based) in the order produced by <see cref="DocumentReader"/>.
    /// </summary>
    public int? BlockIndex { get; set; }

    /// <summary>
    /// Optional source block index within the input document (producer-defined).
    /// For Word, this is the first block index included in the emitted chunk.
    /// </summary>
    public int? SourceBlockIndex { get; set; }

    /// <summary>
    /// Optional 1-based source start line number when the reader can map it accurately.
    /// </summary>
    public int? StartLine { get; set; }

    /// <summary>
    /// Optional 1-based source end line number when the reader can map it accurately.
    /// </summary>
    public int? EndLine { get; set; }

    /// <summary>
    /// Optional 1-based normalized markdown start line used for parser-aware markdown provenance.
    /// </summary>
    public int? NormalizedStartLine { get; set; }

    /// <summary>
    /// Optional 1-based normalized markdown end line used for parser-aware markdown provenance.
    /// </summary>
    public int? NormalizedEndLine { get; set; }

    /// <summary>
    /// Optional heading path label (for example "H1 &gt; H2").
    /// </summary>
    public string? HeadingPath { get; set; }

    /// <summary>
    /// Optional unique slug/anchor for the active heading section.
    /// </summary>
    public string? HeadingSlug { get; set; }

    /// <summary>
    /// Optional source block kind label (for example "heading", "paragraph", "code").
    /// </summary>
    public string? SourceBlockKind { get; set; }

    /// <summary>
    /// Optional deterministic anchor for the first logical block included in the chunk.
    /// For markdown this can identify a heading section or a sub-block within that section.
    /// </summary>
    public string? BlockAnchor { get; set; }

    /// <summary>
    /// Optional sheet name (Excel).
    /// </summary>
    public string? Sheet { get; set; }

    /// <summary>
    /// Optional A1 range descriptor (Excel).
    /// </summary>
    public string? A1Range { get; set; }

    /// <summary>
    /// Optional 1-based slide number (PowerPoint).
    /// </summary>
    public int? Slide { get; set; }

    /// <summary>
    /// Optional 1-based page number (PDF).
    /// </summary>
    public int? Page { get; set; }

    /// <summary>
    /// Optional 0-based table index within the closest source container, such as a PDF page or spreadsheet sheet.
    /// </summary>
    public int? TableIndex { get; set; }
}
