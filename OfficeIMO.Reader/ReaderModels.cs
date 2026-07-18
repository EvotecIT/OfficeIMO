using OfficeIMO.Markdown;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace OfficeIMO.Reader;

/// <summary>
/// The detected input kind used by <see cref="OfficeDocumentReader"/>.
/// </summary>
public enum ReaderInputKind {
    /// <summary>
    /// Unknown/other file kind treated as plain text.
    /// </summary>
    Unknown = 0,
    /// <summary>
    /// Word document (DOCX/DOCM/DOC).
    /// </summary>
    Word = 1,
    /// <summary>
    /// Excel workbook (XLSX/XLSM/XLS).
    /// </summary>
    Excel = 2,
    /// <summary>
    /// PowerPoint presentation (PPTX/PPTM/PPT/POT/PPS).
    /// </summary>
    PowerPoint = 3,
    /// <summary>
    /// Markdown text file.
    /// </summary>
    Markdown = 4,
    /// <summary>
    /// Generic text file.
    /// </summary>
    Text = 5,
    /// <summary>
    /// PDF document.
    /// </summary>
    Pdf = 6,
    /// <summary>
    /// CSV/TSV structured text.
    /// </summary>
    Csv = 7,
    /// <summary>
    /// JSON structured text.
    /// </summary>
    Json = 8,
    /// <summary>
    /// XML structured text.
    /// </summary>
    Xml = 9,
    /// <summary>
    /// HTML document.
    /// </summary>
    Html = 10,
    /// <summary>
    /// ZIP archive.
    /// </summary>
    Zip = 11,
    /// <summary>
    /// EPUB e-book archive.
    /// </summary>
    Epub = 12,
    /// <summary>
    /// Visio drawing (VSDX/VSDM/VSTX/VSTM).
    /// </summary>
    Visio = 13,
    /// <summary>
    /// YAML structured text.
    /// </summary>
    Yaml = 14,
    /// <summary>
    /// Rich Text Format document.
    /// </summary>
    Rtf = 15,
    /// <summary>
    /// OpenDocument Text, Spreadsheet, or Presentation package.
    /// </summary>
    OpenDocument = 16,
    /// <summary>
    /// AsciiDoc technical document.
    /// </summary>
    AsciiDoc = 17,
    /// <summary>
    /// Bounded-profile LaTeX technical document.
    /// </summary>
    Latex = 18,
    /// <summary>
    /// Email, Outlook item, TNEF payload, or mbox mailbox.
    /// </summary>
    Email = 19,
    /// <summary>
    /// Offline Microsoft OneNote artifact.
    /// </summary>
    OneNote = 20
}

/// <summary>
/// A normalized extraction chunk produced by <see cref="OfficeDocumentReader"/>.
/// </summary>
public sealed class ReaderChunk {
    // Adapter projections can span multiple bounded chunks while remaining one
    // logical Markdown block. This is intentionally an internal aggregation
    // contract rather than part of the versioned transport schema.
    internal bool ContinuesPreviousChunk { get; set; }

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

    /// <summary>Number of logical tables whose confidence is below the reader's high-confidence threshold.</summary>
    public int LowConfidenceTableCount { get; set; }

    /// <summary>Number of numeric-looking table columns in the selected scope.</summary>
    public int NumericTableColumnCount { get; set; }

    /// <summary>Number of inferred fallback table column names in the selected scope.</summary>
    public int FallbackTableColumnNameCount { get; set; }

    /// <summary>Number of expected table cells that were empty or unavailable in the selected scope.</summary>
    public int MissingTableCellCount { get; set; }

    /// <summary>Number of logical images in the selected scope.</summary>
    public int ImageCount { get; set; }

    /// <summary>Number of logical images in the selected scope that expose placement geometry.</summary>
    public int ImageGeometryCount { get; set; }

    /// <summary>Ratio of images with placement geometry to all images in the selected scope.</summary>
    public double ImageGeometryCoverage { get; set; }

    /// <summary>Number of logical images in the selected scope whose primary placement is not axis-aligned.</summary>
    public int ImageNonAxisAlignedCount { get; set; }

    /// <summary>Ratio of non-axis-aligned image placements to all images with placement geometry in the selected scope.</summary>
    public double ImageNonAxisAlignedCoverage { get; set; }

    /// <summary>Number of logical link annotations in the selected scope.</summary>
    public int LinkCount { get; set; }

    /// <summary>True when the source exposes readable XMP metadata.</summary>
    public bool HasXmpMetadata { get; set; }

    /// <summary>Number of readable output intents in the source document catalog.</summary>
    public int OutputIntentCount { get; set; }

    /// <summary>Number of readable embedded or associated file attachments in the source document.</summary>
    public int AttachmentCount { get; set; }

    /// <summary>True when the source exposes tagged PDF structure metadata.</summary>
    public bool HasTaggedContent { get; set; }

    /// <summary>Number of readable tagged PDF structure elements in the source document.</summary>
    public int TaggedStructureElementCount { get; set; }

    /// <summary>Number of readable tagged PDF marked-content references in the source document.</summary>
    public int TaggedMarkedContentReferenceCount { get; set; }

    /// <summary>Number of readable optional-content groups/layers in the source document catalog.</summary>
    public int OptionalContentGroupCount { get; set; }

    /// <summary>Number of optional-content groups/layers initially hidden by the default configuration.</summary>
    public int OptionalContentInitiallyHiddenCount { get; set; }

    /// <summary>Number of optional-content groups/layers locked by the default configuration.</summary>
    public int OptionalContentLockedCount { get; set; }

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

    /// <summary>Number of selected actions whose type can execute script, launch external content, submit/import data, or play rich media.</summary>
    public int PotentiallyUnsafeActionCount { get; set; }

    /// <summary>Number of selected JavaScript actions.</summary>
    public int JavaScriptActionCount { get; set; }

    /// <summary>Number of selected launch actions.</summary>
    public int LaunchActionCount { get; set; }

    /// <summary>Number of selected form submission actions.</summary>
    public int SubmitFormActionCount { get; set; }

    /// <summary>Number of selected import-data actions.</summary>
    public int ImportDataActionCount { get; set; }

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
    /// Optional emitted chunk index (0-based) in the order produced by <see cref="OfficeDocumentReader"/>.
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
    /// Escaped heading path used by hierarchy projections. Normal Reader output keeps
    /// <see cref="HeadingPath"/> as the original display value. This value is transported so a
    /// serialized read result can reproduce the same hierarchy as the in-memory result.
    /// </summary>
    public string? HierarchyHeadingPath { get; set; }

    /// <summary>Display-path snapshot associated with <see cref="HierarchyHeadingPath"/>.</summary>
    internal string? HierarchyHeadingDisplayPath { get; set; }

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
