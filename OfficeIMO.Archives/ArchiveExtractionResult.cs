namespace OfficeIMO.Archives;

/// <summary>The container format detected from package bytes.</summary>
public enum ArchiveFormat {
    /// <summary>A Microsoft Cabinet archive.</summary>
    Cab,
    /// <summary>A Windows Installer database with embedded cabinets.</summary>
    Msi,
    /// <summary>A single-disk ZIP archive with stored or Deflate members.</summary>
    Zip,
    /// <summary>A POSIX ustar, pax, GNU or v7 TAR archive.</summary>
    Tar,
    /// <summary>A TAR archive compressed with gzip (.tar.gz, .tgz).</summary>
    TarGZip,
    /// <summary>A single file compressed with gzip (.gz) whose content is not a TAR archive.</summary>
    GZip
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
    internal ArchiveExtractionResult(ArchiveFormat format, IReadOnlyList<ExtractedArchiveFile> files, long totalBytes, ArchiveResourceUsage usage) {
        Format = format; Files = files; TotalBytes = totalBytes; ResourceUsage = usage;
    }
    /// <summary>Cumulative resource charges of the operation at publication.</summary>
    public ArchiveResourceUsage ResourceUsage { get; }
    /// <summary>The detected input format.</summary>
    public ArchiveFormat Format { get; }
    /// <summary>Selected files published into the destination directory.</summary>
    public IReadOnlyList<ExtractedArchiveFile> Files { get; }
    /// <summary>Aggregate complete length of published files.</summary>
    public long TotalBytes { get; }
}
