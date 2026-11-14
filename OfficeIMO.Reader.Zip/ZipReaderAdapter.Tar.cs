using System.Globalization;
using OfficeIMO.Archives;

namespace OfficeIMO.Reader.Zip;

internal static partial class ZipReaderAdapter {
    /// <summary>Whether a seekable source is a TAR, .tar.gz or single-file gzip package rather than a ZIP.</summary>
    private static bool IsTarOrGZip(Stream source, CancellationToken token) {
        try {
            ArchiveFormat format = ArchiveExtractor.Detect(source, token);
            return format == ArchiveFormat.Tar || format == ArchiveFormat.TarGZip || format == ArchiveFormat.GZip;
        } catch (NotSupportedException) {
            return false;
        }
    }

    /// <summary>Reads the supported documents of a TAR, .tar.gz or single-file gzip. Members are verified and
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
            if (ArchiveExtractor.Detect(source, token) == ArchiveFormat.GZip) {
                using var output = new MemoryStream();
                GZipReadResult result = GZipArchive.Decompress(source, output, options, token);
                string name = !string.IsNullOrEmpty(result.FileName) ? Path.GetFileName(result.FileName) : InnerGZipName(archivePath);
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

    private static string InnerGZipName(string archivePath) {
        string name = Path.GetFileName(archivePath.Replace('\\', '/').Split(new[] { "::" }, StringSplitOptions.None).Last());
        return name.EndsWith(".gz", StringComparison.OrdinalIgnoreCase) && name.Length > 3 ? name.Substring(0, name.Length - 3) : "content";
    }

    /// <summary>Returns false for a member over MaxInputBytes (reported as a warning); otherwise reserves Reader capacity for it.</summary>
    private static bool Admit(long length, ReaderOptions readerOptions) {
        if (readerOptions.MaxInputBytes.HasValue && length > readerOptions.MaxInputBytes.Value) return false;
        ReaderNestedContent.ReserveDecodedInput(length);
        return true;
    }
}
