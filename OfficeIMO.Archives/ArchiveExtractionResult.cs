using System.Collections.Generic;

namespace OfficeIMO.Archives;

/// <summary>The container format detected from package bytes.</summary>
public enum ArchiveFormat {
    /// <summary>A Microsoft Cabinet archive.</summary>
    Cab,
    /// <summary>A Windows Installer database with embedded cabinets.</summary>
    Msi
}

/// <summary>A verified file published by an extraction operation.</summary>
public sealed class ExtractedArchiveFile {
    internal ExtractedArchiveFile(string relativePath, long length) { RelativePath = relativePath; Length = length; }
    /// <summary>Portable relative path, using forward slashes and preserving case.
    /// MSI uses installed filename spelling; equivalent aliases use the ordinal-smallest normalized path in their group.</summary>
    public string RelativePath { get; }
    /// <summary>Complete extracted file length in bytes.</summary>
    public long Length { get; }
}

/// <summary>The complete verified result of a single-container extraction.</summary>
public sealed class ArchiveExtractionResult {
    internal ArchiveExtractionResult(ArchiveFormat format, IReadOnlyList<ExtractedArchiveFile> files, long totalBytes) {
        Format = format; Files = files; TotalBytes = totalBytes;
    }
    /// <summary>The detected input format.</summary>
    public ArchiveFormat Format { get; }
    /// <summary>Selected files published into the destination directory.</summary>
    public IReadOnlyList<ExtractedArchiveFile> Files { get; }
    /// <summary>Aggregate complete length of published files.</summary>
    public long TotalBytes { get; }
}
