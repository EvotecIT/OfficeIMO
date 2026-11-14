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
        IReadOnlyList<ArchiveEntryData> members;
        source.Position = 0;
        try {
            if (ArchiveExtractor.Detect(source, token) == ArchiveFormat.GZip) {
                using var output = new MemoryStream();
                GZipReadResult result = GZipArchive.Decompress(source, output, options, token);
                string name = !string.IsNullOrEmpty(result.FileName) ? Path.GetFileName(result.FileName) : InnerGZipName(archivePath);
                members = ShouldAttemptRead(name) ? new[] { new ArchiveEntryData(name, output.ToArray()) } : Array.Empty<ArchiveEntryData>();
            } else {
                members = ArchiveReader.ReadTar(source, options, token);
            }
        } catch (OfficeIMO.OfficeInspectionResourceLimitException error) {
            throw new ReaderResourceLimitException(error.LimitName, error.Maximum, error);
        }

        foreach (ArchiveEntryData member in members) {
            token.ThrowIfCancellationRequested();
            string entryName = member.RelativePath;
            if (readerOptions.MaxInputBytes.HasValue && member.Length > readerOptions.MaxInputBytes.Value) {
                yield return BuildWarningChunk(archiveSource, archivePath, entryName, warningCounter.Next(),
                    "Skipped archive entry because it exceeds MaxInputBytes (" + member.Length.ToString(CultureInfo.InvariantCulture)
                    + " > " + readerOptions.MaxInputBytes.Value.ToString(CultureInfo.InvariantCulture) + ").",
                    readerOptions.ComputeHashes, sourceLengthBytes: member.Length);
                continue;
            }
            string virtualPath = BuildVirtualPath(archivePath, entryName);
            OfficeDocumentReadResult? document = null;
            string? parseError = null;
            try {
                ReaderNestedContent.ReserveDecodedInput(member.Length);
                ReaderNestedContent.MarkReservedPayload(member.Data);
                using var content = new MemoryStream(member.Data, 0, member.Data.Length, writable: false, publiclyVisible: true);
                document = ReaderNestedContent.ReadDocumentInContainer(content, entryName, virtualPath, readerOptions, token);
            } catch (Exception ex) when (ex is not OperationCanceledException and not ReaderResourceLimitException
                and not ReaderAdmissionLimitException and not OfficeIMO.OfficeInspectionResourceLimitException) {
                parseError = "Skipped archive entry due parse error: " + ex.GetType().Name + ".";
            }
            if (parseError != null) {
                yield return BuildWarningChunk(archiveSource, archivePath, entryName, warningCounter.Next(), parseError,
                    readerOptions.ComputeHashes, sourceLengthBytes: member.Length);
                continue;
            }
            DocumentReaderEngine.ApplyExternalSourceMetadata(document!, BuildSourceId(virtualPath), null, member.Length, readerOptions.ComputeHashes);
            foreach (ReaderChunk child in document!.Chunks) {
                token.ThrowIfCancellationRequested();
                ReaderChunk chunk = child.CopyForContainer();
                ApplyVirtualSourceMetadata(chunk, archiveSource, virtualPath, member.Length, default, readerOptions.ComputeHashes);
                yield return chunk;
            }
        }
    }

    private static string InnerGZipName(string archivePath) {
        string name = Path.GetFileName(archivePath.Replace('\\', '/').Split(new[] { "::" }, StringSplitOptions.None).Last());
        return name.EndsWith(".gz", StringComparison.OrdinalIgnoreCase) && name.Length > 3 ? name.Substring(0, name.Length - 3) : "content";
    }
}
