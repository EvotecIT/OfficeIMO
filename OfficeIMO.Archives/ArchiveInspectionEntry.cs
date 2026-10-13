namespace OfficeIMO.Archives;

/// <summary>The declared CAB folder compression, without implying that its payload has been decoded.</summary>
public enum ArchiveCompression {
    /// <summary>No CAB folder compression applies, for example for external MSI source files.</summary>
    None,
    /// <summary>Uncompressed CAB folder bytes.</summary>
    Store,
    /// <summary>MSZIP compressed CAB folder bytes.</summary>
    MsZip,
    /// <summary>LZX compressed CAB folder bytes.</summary>
    Lzx,
    /// <summary>Quantum compression, which is inventoried but cannot be decoded.</summary>
    Quantum,
    /// <summary>An unrecognized or unsupported compression method.</summary>
    Unknown
}

/// <summary>A passive inventory diagnostic. Inventory support does not verify payload integrity.</summary>
public sealed class ArchiveInspectionDiagnostic {
    internal ArchiveInspectionDiagnostic(string code, string message) { Code = code; Message = message; }
    /// <summary>Stable diagnostic identifier.</summary>
    public string Code { get; }
    /// <summary>The unsupported feature or consistency issue.</summary>
    public string Message { get; }
}

/// <summary>One authored CAB file record or MSI File row. Names and aliases are never coalesced.</summary>
public sealed class ArchiveInspectionEntry {
    internal ArchiveInspectionEntry(int ordinal, string rawName, string? rawFileName, string? authoredPath,
        long length, string? cabinetName, int? cabinetOrdinal, int? sequence,
        ArchiveCompression compression, ushort? rawCompression, int? folderOrdinal,
        long? folderCompressedBytes, long? folderExpandedBytes, int? folderFrameCount, int? folderChecksumFrameCount,
        ArchiveInspectionDiagnostic? diagnostic) {
        Ordinal = ordinal; RawName = rawName; RawFileName = rawFileName; AuthoredPath = authoredPath;
        Length = length; CabinetName = cabinetName; CabinetOrdinal = cabinetOrdinal; Sequence = sequence;
        Compression = compression; RawCompression = rawCompression; FolderOrdinal = folderOrdinal;
        FolderCompressedBytes = folderCompressedBytes; FolderExpandedBytes = folderExpandedBytes; Diagnostic = diagnostic;
        FolderFrameCount = folderFrameCount; FolderChecksumFrameCount = folderChecksumFrameCount;
    }
    /// <summary>Zero-based CAB CFFILE order or MSI File-table row order, used by OpenRead.</summary>
    public int Ordinal { get; }
    /// <summary>Exact decoded CAB member name or MSI File identifier, without path normalization.</summary>
    public string RawName { get; }
    /// <summary>Exact MSI FileName field, including any authored short/long name pair; null for CAB.</summary>
    public string? RawFileName { get; }
    /// <summary>Assembled authored MSI target path, without evaluating installation properties or conditions; null for CAB.</summary>
    public string? AuthoredPath { get; }
    /// <summary>Complete declared member size. It is verified again before a member stream is returned.</summary>
    public long Length { get; }
    /// <summary>Exact MSI Media cabinet reference, including '#' for embedded cabinets; null for a standalone CAB.</summary>
    public string? CabinetName { get; }
    /// <summary>CFFILE ordinal within the embedded cabinet; null when no supported embedded source exists.</summary>
    public int? CabinetOrdinal { get; }
    /// <summary>Authored MSI File.Sequence; null for CAB.</summary>
    public int? Sequence { get; }
    /// <summary>Declared folder compression.</summary>
    public ArchiveCompression Compression { get; }
    /// <summary>Raw CAB typeCompress field, including codec parameters; null when no CAB folder exists.</summary>
    public ushort? RawCompression { get; }
    /// <summary>Zero-based folder ordinal in the source cabinet.</summary>
    public int? FolderOrdinal { get; }
    /// <summary>Aggregate compressed CFDATA payload bytes for the shared folder, not a per-file compressed size.</summary>
    public long? FolderCompressedBytes { get; }
    /// <summary>Aggregate declared expansion of the shared folder, including gaps and other members.</summary>
    public long? FolderExpandedBytes { get; }
    /// <summary>Number of CFDATA frames in the complete containing folder.</summary>
    public int? FolderFrameCount { get; }
    /// <summary>Number of containing-folder frames declaring a nonzero CAB checksum. Inventory does not check these values.
    /// A successful OpenRead or ReadMembers verifies these frame checksums; zero-checksum frames have no checksum verification.
    /// CAB frame checksums are corruption checks, not cryptographic hashes or authenticity evidence.</summary>
    public int? FolderChecksumFrameCount { get; }
    /// <summary>True when the declared source and codec can be read. Checksums and decoded integrity are verified by OpenRead.</summary>
    public bool CanRead => Diagnostic == null;
    /// <summary>The reason the member cannot be read through this managed inspection profile.</summary>
    public ArchiveInspectionDiagnostic? Diagnostic { get; }
}

/// <summary>One inert raw stream in an ordinary MSI compound database, including metadata, Binary content and cabinets.</summary>
public sealed class ArchiveStorageStream {
    internal ArchiveStorageStream(int ordinal, int directoryOrder, string rawName, long length) {
        Ordinal = ordinal; DirectoryOrder = directoryOrder; RawName = rawName; Length = length;
    }
    /// <summary>Zero-based index in StorageStreams, used by OpenStorageStream.</summary>
    public int Ordinal { get; }
    /// <summary>Source compound directory traversal order.</summary>
    public int DirectoryOrder { get; }
    /// <summary>Exact compound stream name, retaining MSI's encoded stream-name characters.</summary>
    public string RawName { get; }
    /// <summary>Declared complete raw stream size.</summary>
    public long Length { get; }
}
