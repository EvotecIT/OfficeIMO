using System.IO.Compression;

namespace OfficeIMO.Excel;

/// <summary>Writes the parts of a new package one after another; dispose writes the central directory.</summary>
/// <remarks>Package writers share this surface whether the bytes go through <see cref="ZipArchive"/> or through
/// the owned <c>ExcelZipWriter</c> (.NET 8 or later), which handles <see cref="CompressionLevel.Fastest"/> with its own deflate encoder.</remarks>
internal abstract class ExcelPackageArchive : IDisposable {
    /// <summary>Creates a package over <paramref name="output"/>, using the owned writer where it is available.</summary>
    /// <param name="output">The destination; it is not disposed.</param>
    /// <param name="modified">A fixed entry timestamp, or <see langword="null"/> for the time the package is created.</param>
    internal static ExcelPackageArchive Create(Stream output, DateTime? modified = null) {
#if NET8_0_OR_GREATER
        if (BitConverter.IsLittleEndian) return new ExcelZipWriter(output, modified ?? DateTime.Now);
#endif
        return new PlatformArchive(output, modified);
    }

    /// <summary>Opens the next entry; dispose the returned stream before opening another.</summary>
    internal abstract Stream OpenEntry(string name, CompressionLevel level);

    internal void WriteEntry(string name, byte[] content, CompressionLevel level) =>
        WriteEntry(name, new ArraySegment<byte>(content), level);

    internal void WriteEntry(string name, ArraySegment<byte> content, CompressionLevel level) {
        using Stream entry = OpenEntry(name, level);
        entry.Write(content.Array!, content.Offset, content.Count);
    }

    internal void WriteEntry(string name, string content, CompressionLevel level) =>
        WriteEntry(name, new System.Text.UTF8Encoding(encoderShouldEmitUTF8Identifier: false).GetBytes(content), level);

    private WorksheetCompressionPipeline? _compression;

    /// <summary>
    /// Returns the package's background compressor aimed at <paramref name="destination"/>. Entries are written one
    /// after another, so one worker and its buffers serve every sheet; each entry drains it before closing.
    /// </summary>
    internal WorksheetCompressionPipeline AcquireCompression(Stream destination) {
        if (_compression == null) _compression = new WorksheetCompressionPipeline(destination);
        else _compression.Retarget(destination);
        return _compression;
    }

    /// <summary>Stops the compressor after an entry failed; the next entry starts a new one.</summary>
    internal void AbandonCompression() {
        WorksheetCompressionPipeline? compression = _compression;
        _compression = null;
        compression?.Dispose();
    }

    public abstract void Dispose();

    private sealed class PlatformArchive : ExcelPackageArchive {
        private readonly ZipArchive _archive;
        private readonly DateTimeOffset? _modified;

        internal PlatformArchive(Stream output, DateTime? modified) {
            // .NET Framework's ZipArchive reads Position even when the destination is forward-only.
            Stream destination = output.CanSeek || output is ExcelPositionReportingWriteStream ? output : new ExcelPositionReportingWriteStream(output);
            _archive = new ZipArchive(destination, ZipArchiveMode.Create, leaveOpen: true);
            _modified = modified.HasValue ? new DateTimeOffset(DateTime.SpecifyKind(modified.Value, DateTimeKind.Unspecified), TimeSpan.Zero) : null;
        }

        internal override Stream OpenEntry(string name, CompressionLevel level) {
            ZipArchiveEntry entry = _archive.CreateEntry(name, level);
            if (_modified.HasValue) entry.LastWriteTime = _modified.Value;
            return entry.Open();
        }

        public override void Dispose() {
            AbandonCompression();
            _archive.Dispose();
        }
    }
}
