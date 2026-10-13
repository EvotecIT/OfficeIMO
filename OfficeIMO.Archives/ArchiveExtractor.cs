using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using OfficeIMO.Core.Internal.Cabinet;

namespace OfficeIMO.Archives;

/// <summary>Extracts one managed archive container into a fresh directory without executing its contents.</summary>
public static partial class ArchiveExtractor {
    /// <summary>Verifies and extracts selected files from a CAB or ordinary MSI package.
    /// The destination must be absent or empty. A completed staging directory is published by a same-volume move.</summary>
    /// <exception cref="OfficeInspectionResourceLimitException">A configured or fixed managed admission ceiling is exhausted.</exception>
    /// <exception cref="InvalidDataException">The input is corrupt, unsafe or conflicting.</exception>
    /// <exception cref="NotSupportedException">The input uses an unsupported format or extraction feature.</exception>
    /// <exception cref="IOException">The destination is occupied, contains links, or cannot be published.</exception>
    public static ArchiveExtractionResult Extract(string packagePath, string destinationDirectory,
        ArchiveExtractionOptions? options = null, CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(packagePath)) throw new ArgumentException("A package path is required.", nameof(packagePath));
        if (string.IsNullOrWhiteSpace(destinationDirectory)) throw new ArgumentException("A destination directory is required.", nameof(destinationDirectory));
        var limits = new ArchiveLimits(options ?? new ArchiveExtractionOptions());
        cancellationToken.ThrowIfCancellationRequested();
        string package = Path.GetFullPath(packagePath);
        string destination = Path.GetFullPath(destinationDirectory).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        string parent = Path.GetDirectoryName(destination) ?? throw new ArgumentException("The destination must have a parent directory.", nameof(destinationDirectory));
        ArchivePaths.RejectLinks(destination);
        CheckDestination(destination);
        using var source = new FileStream(package, FileMode.Open, FileAccess.Read, FileShare.Read);
        ArchiveLimits.Check(source.Length, limits.Input, nameof(ArchiveExtractionOptions.MaxInputBytes));
        ArchiveFormat format = Detect(source, cancellationToken);
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
                default: throw new NotSupportedException("Unsupported archive format.");
            }
            cancellationToken.ThrowIfCancellationRequested();
            ArchivePaths.RejectLinks(destination);
            CheckDestination(destination);
            if (Directory.Exists(destination)) Directory.Delete(destination, recursive: false);
            Directory.Move(staging, destination);
            staging = string.Empty;
            return new ArchiveExtractionResult(format, Array.AsReadOnly(files.OrderBy(file => file.RelativePath, StringComparer.Ordinal).ToArray()), total);
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
    /// <summary>Detects a seekable container from its initial header and restores the caller's stream position.</summary>
    internal static ArchiveFormat Detect(Stream source, CancellationToken token) {
        long originalPosition = source.Position;
        try {
            source.Position = 0;
            var header = new byte[8];
            int headerLength = 0;
            while (headerLength < header.Length) {
                token.ThrowIfCancellationRequested();
                int read = source.Read(header, headerLength, header.Length - headerLength);
                if (read == 0) break;
                headerLength += read;
            }
            return Detect(header, headerLength);
        } finally {
            source.Position = originalPosition;
        }
    }
    private static ArchiveFormat Detect(byte[] bytes, int length) {
        if (length >= 4 && bytes[0] == 'M' && bytes[1] == 'S' && bytes[2] == 'C' && bytes[3] == 'F') return ArchiveFormat.Cab;
        byte[] compound = { 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1 };
        if (length == 8 && bytes.SequenceEqual(compound)) return ArchiveFormat.Msi;
        throw new NotSupportedException("The package is not a supported Cabinet or Windows Installer container.");
    }
    private static void ExtractCab(Stream source, string staging, ArchiveLimits limits, List<ExtractedArchiveFile> files,
        ref long total, CancellationToken token) {
        var paths = new ArchivePaths(limits);
        var selected = new Dictionary<string, string>(StringComparer.Ordinal);
        IReadOnlyList<OfficeCabinetEntry> entries = OfficeCabinetArchiveReader.Read(source, limits.Input, limits.Expanded,
            limits.Entry, limits.Entries, token, name => {
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
        string target = Path.Combine(staging, path.Replace('/', Path.DirectorySeparatorChar));
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
