using System.Globalization;
using OfficeIMO.Archives;

namespace OfficeIMO.Reader.Zip;

internal static partial class ZipReaderAdapter {
    /// <summary>Whether a seekable source is a TAR, plain or compressed, or a single gzip, Zstandard, xz or lzip file
    /// rather than a ZIP.</summary>
    private static bool IsTarOrGZip(Stream source, CancellationToken token) {
        try {
            ArchiveFormat format = ArchiveExtractor.Detect(source, token);
            return format == ArchiveFormat.Tar || StreamCodecs.Of(format) != StreamCodec.None;
        } catch (NotSupportedException) {
            return false;
        }
    }

    /// <summary>Reads the supported documents of a TAR (plain, .tar.gz, .tar.zst, .tar.xz or .tar.lz) or of a single gzip,
    /// Zstandard, xz or lzip file. Members are verified and
    /// path-checked by OfficeIMO.Archives; links, devices and directories are never read. Nested archives are not traversed.</summary>
    private static IEnumerable<ReaderChunk> ReadTarOrGZip(Stream source, ArchiveSourceMetadata archiveSource, string archivePath,
        ReaderOptions readerOptions, ArchiveTraversalOptions zipOptions, WarningCounter warningCounter, CancellationToken token) {
        var options = new ArchiveReadOptions {
            OperationBudget = zipOptions.OperationBudget ?? new ArchiveOperationBudget(),
            MaxInputBytes = Math.Max(1, zipOptions.MaxArchiveBytes),
            MaxEntryCount = Math.Max(1, zipOptions.MaxPhysicalEntries),
            MaxEntryBytes = Math.Max(1, zipOptions.MaxEntryUncompressedBytes ?? long.MaxValue),
            MaxExpandedBytes = Math.Max(1, zipOptions.MaxTotalUncompressedBytes ?? long.MaxValue),
            IncludeFile = ShouldAttemptRead
        };
        IReadOnlyList<(string Path, long Length, byte[]? Data)> members;
        source.Position = 0;
        try {
            ArchiveFormat format = ArchiveExtractor.Detect(source, token);
            if (format is ArchiveFormat.GZip or ArchiveFormat.Zstandard or ArchiveFormat.Xz or ArchiveFormat.Lzip) {
                using var output = new MemoryStream();
                string? storedName = null;
                switch (format) {
                    case ArchiveFormat.GZip: storedName = GZipArchive.Decompress(source, output, options, token).FileName; break;
                    case ArchiveFormat.Zstandard: ZstandardArchive.Decompress(source, output, options, token); break;
                    case ArchiveFormat.Xz: XzArchive.Decompress(source, output, options, token); break;
                    default: LzipArchive.Decompress(source, output, options, token); break;
                }
                string name = !string.IsNullOrEmpty(storedName) ? Path.GetFileName(storedName) : InnerName(archivePath, StreamCodecs.Suffix(StreamCodecs.Of(format)));
                byte[] data = output.ToArray();
                members = !ShouldAttemptRead(name) ? Array.Empty<(string, long, byte[]?)>()
                    : new[] { (name, data.LongLength, Admit(data.LongLength, readerOptions) ? data : null) };
            } else {
                // Reader capacity is reserved before each member's payload is allocated.
                members = ArchiveReader.ReadTarMembers(source, options, (_, length) => Admit(length, readerOptions), token);
            }
        } catch (OfficeIMO.OfficeInspectionResourceLimitException error) {
            throw new ReaderResourceLimitException(error.LimitName, error.Maximum, error);
        }

        foreach (var (entryName, length, data) in members) {
            token.ThrowIfCancellationRequested();
            if (data == null) {
                yield return BuildWarningChunk(archiveSource, archivePath, entryName, warningCounter.Next(),
                    "Skipped archive entry because it exceeds MaxInputBytes (" + length.ToString(CultureInfo.InvariantCulture)
                    + " > " + readerOptions.MaxInputBytes!.Value.ToString(CultureInfo.InvariantCulture) + ").",
                    readerOptions.ComputeHashes, sourceLengthBytes: length);
                continue;
            }
            string virtualPath = BuildVirtualPath(archivePath, entryName);
            OfficeDocumentReadResult? document = null;
            string? parseError = null;
            try {
                ReaderNestedContent.MarkReservedPayload(data);
                using var content = new MemoryStream(data, 0, data.Length, writable: false, publiclyVisible: true);
                document = ReaderNestedContent.ReadDocumentInContainer(content, entryName, virtualPath, readerOptions, token);
            } catch (Exception ex) when (ex is not OperationCanceledException and not ReaderResourceLimitException
                and not ReaderAdmissionLimitException and not OfficeIMO.OfficeInspectionResourceLimitException) {
                parseError = "Skipped archive entry due parse error: " + ex.GetType().Name + ".";
            }
            if (parseError != null) {
                yield return BuildWarningChunk(archiveSource, archivePath, entryName, warningCounter.Next(), parseError,
                    readerOptions.ComputeHashes, sourceLengthBytes: length);
                continue;
            }
            DocumentReaderEngine.ApplyExternalSourceMetadata(document!, BuildSourceId(virtualPath), null, length, readerOptions.ComputeHashes);
            foreach (ReaderChunk child in document!.Chunks) {
                token.ThrowIfCancellationRequested();
                ReaderChunk chunk = child.CopyForContainer();
                ApplyVirtualSourceMetadata(chunk, archiveSource, virtualPath, length, default, readerOptions.ComputeHashes);
                yield return chunk;
            }
        }
    }

    private static string InnerName(string archivePath, string suffix) {
        string name = Path.GetFileName(archivePath.Replace('\\', '/').Split(new[] { "::" }, StringSplitOptions.None).Last());
        return name.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) && name.Length > suffix.Length ? name.Substring(0, name.Length - suffix.Length) : "content";
    }

    /// <summary>Returns false for a member over MaxInputBytes (reported as a warning); otherwise reserves Reader capacity for it.</summary>
    private static bool Admit(long length, ReaderOptions readerOptions) {
        if (readerOptions.MaxInputBytes.HasValue && length > readerOptions.MaxInputBytes.Value) return false;
        ReaderNestedContent.ReserveDecodedInput(length);
        return true;
    }
}
