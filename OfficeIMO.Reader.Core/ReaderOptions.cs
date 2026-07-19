using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace OfficeIMO.Reader;

/// <summary>
/// Options controlling extraction behavior for <see cref="OfficeDocumentReader"/>.
/// </summary>
public sealed class ReaderOptions {
    internal const long DefaultOpenXmlMaxCharactersInPart = 10_000_000L;
    internal const int DefaultDetectionMaxProbeBytes = 64 * 1024;
    internal const int DefaultDetectionMaxContainerEntries = 512;
    internal const int MaximumDetectionProbeBytes = 4 * 1024 * 1024;
    internal const int MaximumDetectionContainerEntries = 4096;
    internal const int DefaultMaxOpenXmlImageAssets = 512;
    internal const int DefaultMaxOpenXmlImagePlacementsPerRelationship = 256;
    internal const long DefaultMaxOpenXmlImageAssetBytes = 32L * 1024L * 1024L;
    internal const long DefaultMaxOpenXmlImageTotalAssetBytes = 128L * 1024L * 1024L;

    /// <summary>
    /// Optional maximum input size in bytes enforced by <see cref="OfficeDocumentReader"/> when reading from a file or seekable stream.
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
    /// When true, computes source/chunk hashes for incremental indexing workflows. Default: true.
    /// </summary>
    public bool ComputeHashes { get; set; } = true;

    /// <summary>
    /// Controls extension/content detection during reads. The default inspects content only when the extension is unknown.
    /// </summary>
    public ReaderDetectionMode DetectionMode { get; set; } = ReaderDetectionMode.ContentWhenUnknown;

    /// <summary>
    /// Maximum prefix bytes inspected when content detection runs. Default: 64 KiB.
    /// </summary>
    public int DetectionMaxProbeBytes { get; set; } = DefaultDetectionMaxProbeBytes;

    /// <summary>
    /// Maximum ZIP entries inspected when classifying Office, Visio, EPUB, and ZIP containers. Default: 512.
    /// </summary>
    public int DetectionMaxContainerEntries { get; set; } = DefaultDetectionMaxContainerEntries;
}

/// <summary>
/// Options controlling folder enumeration for <see cref="OfficeDocumentReader.ReadFolder"/>.
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
    /// Optional maximum total bytes accepted for parsing across the folder operation.
    /// Files whose known size exceeds the remaining budget are skipped before their format handler runs.
    /// When source size metadata is unavailable, the cap is enforced after parsing as a best-effort fallback.
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
