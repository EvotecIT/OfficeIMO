using OfficeIMO.Markdown;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace OfficeIMO.Reader;

/// <summary>
/// Options controlling extraction behavior for <see cref="DocumentReader"/>.
/// </summary>
public sealed class ReaderOptions {
    internal const long DefaultOpenXmlMaxCharactersInPart = 10_000_000L;
    internal const int DefaultMaxOpenXmlImageAssets = 512;
    internal const int DefaultMaxOpenXmlImagePlacementsPerRelationship = 256;
    internal const long DefaultMaxOpenXmlImageAssetBytes = 32L * 1024L * 1024L;
    internal const long DefaultMaxOpenXmlImageTotalAssetBytes = 128L * 1024L * 1024L;

    /// <summary>
    /// Optional maximum input size in bytes enforced by <see cref="DocumentReader"/> when reading from a file or seekable stream.
    /// When null, no size limit is enforced.
    /// </summary>
    public long? MaxInputBytes { get; set; }

    /// <summary>
    /// OpenXML security: maximum characters allowed per part when opening OpenXML packages (best-effort).
    /// When null, the OpenXML SDK default is used.
    /// </summary>
    public long? OpenXmlMaxCharactersInPart { get; set; } = DefaultOpenXmlMaxCharactersInPart;

    /// <summary>
    /// OpenXML security: maximum image asset entries emitted from an Office package.
    /// When null, no image asset entry limit is enforced.
    /// </summary>
    public int? MaxOpenXmlImageAssets { get; set; } = DefaultMaxOpenXmlImageAssets;

    /// <summary>
    /// Optional password used to open encrypted Office files when the selected reader engine supports decryption.
    /// </summary>
    public string? OpenPassword { get; set; }

    /// <summary>
    /// OpenXML security: maximum placements emitted for one image relationship.
    /// When null, no per-relationship placement limit is enforced.
    /// </summary>
    public int? MaxOpenXmlImagePlacementsPerRelationship { get; set; } = DefaultMaxOpenXmlImagePlacementsPerRelationship;

    /// <summary>
    /// OpenXML security: maximum bytes read from a single image part.
    /// When null, no individual image payload limit is enforced.
    /// </summary>
    public long? MaxOpenXmlImageAssetBytes { get; set; } = DefaultMaxOpenXmlImageAssetBytes;

    /// <summary>
    /// OpenXML security: maximum unique image payload bytes retained while extracting assets.
    /// Repeated placements of the same image part count once toward this limit.
    /// When null, no total image payload limit is enforced.
    /// </summary>
    public long? MaxOpenXmlImageTotalAssetBytes { get; set; } = DefaultMaxOpenXmlImageTotalAssetBytes;

    /// <summary>
    /// Maximum characters per emitted chunk (best-effort).
    /// </summary>
    public int MaxChars { get; set; } = 8_000;

    /// <summary>
    /// Maximum number of table rows included per table chunk (best-effort).
    /// </summary>
    public int MaxTableRows { get; set; } = 200;

    /// <summary>
    /// When true, include Word footnotes as a final chunk. Default: true.
    /// </summary>
    public bool IncludeWordFootnotes { get; set; } = true;

    /// <summary>
    /// When true, include PowerPoint speaker notes when present. Default: true.
    /// </summary>
    public bool IncludePowerPointNotes { get; set; } = true;

    /// <summary>
    /// Excel: when true, treat the first row as headers. Default: true.
    /// </summary>
    public bool ExcelHeadersInFirstRow { get; set; } = true;

    /// <summary>
    /// Excel: number of worksheet rows per emitted chunk. Default: 200.
    /// </summary>
    public int ExcelChunkRows { get; set; } = 200;

    /// <summary>
    /// Excel: optional sheet name. When null, all sheets are extracted.
    /// </summary>
    public string? ExcelSheetName { get; set; }

    /// <summary>
    /// Excel: optional A1 range. When null, the sheet's used range is used.
    /// </summary>
    public string? ExcelA1Range { get; set; }

    /// <summary>
    /// Markdown: when true, chunk by headings where possible. Default: true.
    /// </summary>
    public bool MarkdownChunkByHeadings { get; set; } = true;

    /// <summary>
    /// Markdown: optional input normalization applied before parser-aware chunking.
    /// This is intended for compact AI/chat markdown fixes while preserving default strict behavior when null.
    /// </summary>
    public MarkdownInputNormalizationOptions? MarkdownInputNormalization { get; set; }

    /// <summary>
    /// When true, computes source/chunk hashes for incremental indexing workflows. Default: true.
    /// </summary>
    public bool ComputeHashes { get; set; } = true;
}

/// <summary>
/// Options controlling folder enumeration for <see cref="DocumentReader.ReadFolder(string, ReaderFolderOptions?, ReaderOptions?, System.Threading.CancellationToken)"/>.
/// </summary>
public sealed class ReaderFolderOptions {
    /// <summary>
    /// When true, enumerates all subdirectories. Default: true.
    /// </summary>
    public bool Recurse { get; set; } = true;

    /// <summary>
    /// Maximum number of files enumerated. Default: 500.
    /// </summary>
    public int MaxFiles { get; set; } = 500;

    /// <summary>
    /// Optional maximum total bytes across all enumerated files (best-effort).
    /// When null, no cap is enforced.
    /// </summary>
    public long? MaxTotalBytes { get; set; }

    /// <summary>
    /// Optional allowed extensions (lower/upper insensitive). When null, a default set plus
    /// currently registered custom handler extensions is used.
    /// Examples: ".docx", ".xlsx", ".pptx", ".md".
    /// </summary>
    public IReadOnlyList<string>? Extensions { get; set; }

    /// <summary>
    /// When true, directory traversal skips reparse points (junctions/symlinks). Default: true.
    /// </summary>
    public bool SkipReparsePoints { get; set; } = true;

    /// <summary>
    /// When true, folder traversal is deterministic (ordinal path ordering). Default: true.
    /// </summary>
    public bool DeterministicOrder { get; set; } = true;
}
