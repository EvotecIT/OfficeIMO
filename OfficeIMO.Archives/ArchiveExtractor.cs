using OfficeIMO.Core.Internal.Cabinet;

namespace OfficeIMO.Archives;

/// <summary>Extracts one managed archive container into a fresh directory without executing its contents.</summary>
public static partial class ArchiveExtractor {
    /// <summary>Verifies and extracts selected files from a ZIP, CAB, ordinary MSI, TAR, .tar.gz or single-file gzip package.
    /// TAR symbolic links, hard links and device records are never created; selecting one fails the extraction.
    /// A gzip file is extracted under its stored name, or the package name without ".gz".
    /// The destination must be absent or empty. A completed staging directory is published by a same-volume move.</summary>
    /// <exception cref="OfficeInspectionResourceLimitException">A configured or fixed managed admission ceiling is exhausted.</exception>
    /// <exception cref="InvalidDataException">The input is corrupt, unsafe or conflicting.</exception>
    /// <exception cref="NotSupportedException">The input uses an unsupported format or extraction feature.</exception>
    /// <exception cref="IOException">The destination is occupied, contains links, or cannot be published.</exception>
    public static ArchiveExtractionResult Extract(string packagePath, string destinationDirectory,
        ArchiveExtractionOptions? options = null, CancellationToken cancellationToken = default) {
        ArgumentException.ThrowIfNullOrWhiteSpace(packagePath);
        ArgumentException.ThrowIfNullOrWhiteSpace(destinationDirectory);
        var limits = new ArchiveLimits(options ?? new ArchiveExtractionOptions());
        cancellationToken.ThrowIfCancellationRequested();
        string package = Path.GetFullPath(packagePath);
        string destination = Path.GetFullPath(destinationDirectory).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        string parent = Path.GetDirectoryName(destination) ?? throw new ArgumentException("The destination must have a parent directory.", nameof(destinationDirectory));
        ArchivePaths.RejectLinks(destination);
        CheckDestination(destination);
        using var source = new FileStream(package, FileMode.Open, FileAccess.Read, FileShare.Read);
        ArchiveLimits.Check(source.Length, limits.Input, nameof(ArchiveExtractionOptions.MaxInputBytes));
        limits.Operation.AdmitInput(source.Length);
        ArchiveFormat format = Detect(source, cancellationToken);
        if (limits.AllowedFormats != null && !limits.AllowedFormats.Contains(format))
            throw new NotSupportedException("The detected archive format is not allowed by this extraction operation: " + format);
        Directory.CreateDirectory(parent);
        string staging = Path.Combine(parent, ".officeimo-archive-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(staging);
        try {
            ArchivePaths.RejectLinks(staging);
            var files = new List<ExtractedArchiveFile>();
            long total = 0;
            switch (format) {
                case ArchiveFormat.Cab: ExtractCab(source, staging, limits, files, ref total, cancellationToken); break;
                case ArchiveFormat.Msi: ExtractMsi(source, staging, limits, files, ref total, cancellationToken); break;
                case ArchiveFormat.Zip: ExtractZip(source, staging, limits, files, ref total, cancellationToken); break;
                case ArchiveFormat.Tar:
                case ArchiveFormat.TarGZip: ExtractTar(source, staging, limits, files, ref total, cancellationToken); break;
                case ArchiveFormat.GZip: ExtractGZip(source, package, staging, limits, files, ref total, cancellationToken); break;
                default: throw new NotSupportedException("Unsupported archive format.");
            }
            cancellationToken.ThrowIfCancellationRequested();
            ArchivePaths.RejectLinks(destination);
            CheckDestination(destination);
            if (Directory.Exists(destination)) Directory.Delete(destination, recursive: false);
            Directory.Move(staging, destination);
            staging = string.Empty;
            return new ArchiveExtractionResult(format, Array.AsReadOnly(files.OrderBy(file => file.RelativePath, StringComparer.Ordinal).ToArray()), total, limits.Operation.Usage);
        } catch (OfficeCabinetException exception) {
            throw limits.Translate(exception);
        } finally {
            if (staging.Length != 0 && Directory.Exists(staging)) Directory.Delete(staging, recursive: true);
        }
    }
    private static void CheckDestination(string destination) {
        if (File.Exists(destination) || Directory.Exists(destination) && Directory.EnumerateFileSystemEntries(destination).Any())
            throw new IOException("Extraction requires an absent or empty destination directory.");
    }
    /// <summary>Detects a seekable container from its initial bytes and restores the caller's stream position.
    /// Gzip input is classified as TAR when its first decompressed block is a TAR header.</summary>
    internal static ArchiveFormat Detect(Stream source, CancellationToken token) {
        long originalPosition = source.Position;
        try {
            source.Position = 0;
            // Read only the 8-byte magic first; TAR detection needs the whole first header block.
            var header = new byte[1024];
            int headerLength = Fill(source, header, 8, token);
            if (GZipContent.HasSignature(header, headerLength)) {
                source.Position = 0;
                var peek = new byte[1024];
                return LooksLikeTar(peek, GZipContent.Peek(source, peek, peek.Length)) ? ArchiveFormat.TarGZip : ArchiveFormat.GZip;
            }
            if (TryDetectMagic(header, headerLength, out ArchiveFormat format)) return format;
            if (headerLength == 8) headerLength = Fill(source, header, header.Length, token, start: 8);
            if (LooksLikeTar(header, headerLength)) return ArchiveFormat.Tar;
            throw new NotSupportedException("The package is not a supported ZIP, Cabinet, Windows Installer, TAR or gzip container.");
        } finally {
            source.Position = originalPosition;
        }
    }
    private static int Fill(Stream source, byte[] buffer, int count, CancellationToken token, int start = 0) {
        int length = start;
        while (length < count) {
            token.ThrowIfCancellationRequested();
            int read = source.Read(buffer, length, count - length);
            if (read == 0) break;
            length += read;
        }
        return length;
    }
    private static bool TryDetectMagic(byte[] bytes, int length, out ArchiveFormat format) {
        format = ArchiveFormat.Zip;
        if (length >= 4 && bytes[0] == 'M' && bytes[1] == 'S' && bytes[2] == 'C' && bytes[3] == 'F') { format = ArchiveFormat.Cab; return true; }
        byte[] compound = { 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1 };
        if (length >= 8 && bytes.Take(8).SequenceEqual(compound)) { format = ArchiveFormat.Msi; return true; }
        // A local file header, or the directory signature/classic/ZIP64 end record of an empty archive.
        return length >= 4 && bytes[0] == 'P' && bytes[1] == 'K' && (bytes[2] == 3 && bytes[3] == 4
            || bytes[2] == 5 && (bytes[3] == 5 || bytes[3] == 6) || bytes[2] == 6 && bytes[3] == 6);
    }
    /// <summary>A checksum-valid TAR header, or the two zero blocks of an empty TAR archive.</summary>
    internal static bool LooksLikeTar(byte[] bytes, int length) => TarArchiveContent.IsHeader(bytes, length)
        || length >= 1024 && TarArchiveContent.IsZeroBlock(bytes) && TarArchiveContent.IsZeroBlock(bytes, 512);
    private static void ExtractCab(Stream source, string staging, ArchiveLimits limits, List<ExtractedArchiveFile> files,
        ref long total, CancellationToken token) {
        var paths = new ArchivePaths(limits);
        var selected = new Dictionary<string, string>(StringComparer.Ordinal);
        IReadOnlyList<OfficeCabinetEntry> entries = OfficeCabinetArchiveReader.Read(source, limits.Input, limits.Expanded,
            limits.Entry, limits.Entries, token, name => {
                limits.Operation.AdmitEntries(1);
                string path = paths.Add(name);
                if (!limits.Select(path)) return false;
                selected.Add(name, path); return true;
            }, limits.DecompressionWork);
        foreach (OfficeCabinetEntry entry in entries) {
            token.ThrowIfCancellationRequested();
            WriteBytes(staging, selected[entry.Name], entry.Data, limits, files, ref total, token);
        }
    }
    private static void WriteBytes(string staging, string path, byte[] bytes, ArchiveLimits limits,
        List<ExtractedArchiveFile> files, ref long total, CancellationToken token) {
        using var memory = new MemoryStream(bytes, writable: false);
        WriteStream(staging, path, memory, bytes.LongLength, limits, files, ref total, token);
    }
    private static void WriteStream(string staging, string path, Stream payload, long length, ArchiveLimits limits,
        List<ExtractedArchiveFile> files, ref long total, CancellationToken token) {
        limits.CheckEntry(length);
        ArchiveLimits.Add(ref total, length, limits.Expanded, "MaxExpandedBytes");
        string target = Path.GetFullPath(Path.Combine(staging, path.Replace('/', Path.DirectorySeparatorChar)));
        string prefix = Path.GetFullPath(staging).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            + Path.DirectorySeparatorChar;
        if (!target.StartsWith(prefix, Path.DirectorySeparatorChar == '\\' ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal))
            throw new InvalidDataException("Archive member escapes the staging directory.");
        ArchivePaths.RejectLinks(target, staging);
        Directory.CreateDirectory(Path.GetDirectoryName(target)!);
        ArchivePaths.RejectLinks(target, staging);
        using var output = new FileStream(target, FileMode.CreateNew, FileAccess.Write, FileShare.None);
        byte[] buffer = new byte[81920];
        long written = 0;
        while (written < length) {
            token.ThrowIfCancellationRequested();
            int read = payload.Read(buffer, 0, (int)Math.Min(buffer.Length, length - written));
            if (read <= 0) throw new InvalidDataException("Archive member is shorter than its declared length.");
            output.Write(buffer, 0, read);
            written += read;
        }
        token.ThrowIfCancellationRequested();
        if (payload.ReadByte() != -1) throw new InvalidDataException("Archive member exceeds its declared length.");
        files.Add(new ExtractedArchiveFile(path, length));
    }
}
