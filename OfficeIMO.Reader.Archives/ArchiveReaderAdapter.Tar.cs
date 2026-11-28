using OfficeIMO.Archives;

namespace OfficeIMO.Reader.Archives;

internal static partial class ArchiveReaderAdapter {
    /// <summary>Reads a TAR (plain, .tar.gz, .tar.zst, .tar.xz or .tar.lz) or a single gzip, Zstandard, xz or lzip file.
    /// Members are verified and path-checked by OfficeIMO.Archives; links, devices and directories are never read.</summary>
    private static IEnumerable<ReaderChunk> ReadTarOrCompressedFile(Stream source, ArchiveFormat format, string archivePath, int nestedDepth,
        ArchiveOperationBudget budget, ReadContext context) {
        var options = ReadOptions(context, budget, path => IsWanted(path, context.Archive));
        IReadOnlyList<(string Path, long Length, byte[]? Data)> members;
        source.Position = 0;
        try {
            if (format is ArchiveFormat.GZip or ArchiveFormat.Zstandard or ArchiveFormat.Xz or ArchiveFormat.Lzip) {
                using var output = new MemoryStream();
                string? storedName = null;
                switch (format) {
                    case ArchiveFormat.GZip: storedName = GZipArchive.Decompress(source, output, options, context.Token).FileName; break;
                    case ArchiveFormat.Zstandard: ZstandardArchive.Decompress(source, output, options, context.Token); break;
                    case ArchiveFormat.Xz: XzArchive.Decompress(source, output, options, context.Token); break;
                    default: LzipArchive.Decompress(source, output, options, context.Token); break;
                }
                string name = !string.IsNullOrEmpty(storedName) ? Path.GetFileName(storedName) : InnerName(archivePath, StreamCodecs.Suffix(StreamCodecs.Of(format)));
                byte[] data = output.ToArray();
                members = !IsWanted(name, context.Archive) ? Array.Empty<(string, long, byte[]?)>()
                    : new[] { (name, data.LongLength, Admit(name, data.LongLength, nestedDepth, context) ? data : null) };
            } else {
                // Reader capacity is reserved before each member's payload is allocated.
                members = ArchiveReader.ReadTarMembers(source, options, (path, length) => Admit(path, length, nestedDepth, context), context.Token);
            }
        } catch (OfficeIMO.OfficeInspectionResourceLimitException error) {
            throw new ReaderResourceLimitException(error.LimitName, error.Maximum, error);
        }

        foreach (var (entryName, length, data) in members) {
            context.Token.ThrowIfCancellationRequested();
            foreach (var chunk in EmitMember(entryName, archivePath, length, null, () => Admitted(data, entryName), nestedDepth, budget, context))
                yield return chunk;
        }
    }

    /// <summary>Options for the in-memory readers, bounded by the traversal limits and charged to <paramref name="budget"/>.</summary>
    private static ArchiveReadOptions ReadOptions(ReadContext context, ArchiveOperationBudget budget, Func<string, bool>? include) => new() {
        OperationBudget = budget,
        MaxInputBytes = Math.Max(1, context.Traversal.MaxArchiveBytes),
        MaxEntryCount = Math.Max(1, context.Traversal.MaxPhysicalEntries),
        MaxEntryBytes = Math.Max(1, context.Traversal.MaxEntryUncompressedBytes ?? long.MaxValue),
        MaxExpandedBytes = Math.Max(1, context.Traversal.MaxTotalUncompressedBytes ?? long.MaxValue),
        IncludeFile = include
    };

    private static string InnerName(string archivePath, string suffix) {
        string name = Path.GetFileName(archivePath.Replace('\\', '/').Split(new[] { "::" }, StringSplitOptions.None).Last());
        return name.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) && name.Length > suffix.Length ? name.Substring(0, name.Length - suffix.Length) : "content";
    }

    /// <summary>Returns false for a member <see cref="EmitMember"/> will skip with a warning before reading it; otherwise
    /// reserves Reader capacity for it.</summary>
    private static bool Admit(string path, long length, int nestedDepth, ReadContext context) {
        ReaderOptions readerOptions = context.ReaderOptions;
        ReaderArchiveOptions archive = context.Archive;
        if (readerOptions.MaxInputBytes.HasValue && length > readerOptions.MaxInputBytes.Value || length > int.MaxValue) return false;
        if (IsNestedArchive(path, archive) && (!archive.ReadNestedArchives || nestedDepth >= archive.MaxNestedDepth
            || archive.MaxNestedArchiveBytes.HasValue && length > archive.MaxNestedArchiveBytes.Value)) return false;
        ReaderNestedContent.ReserveDecodedInput(length);
        return true;
    }

    private static byte[] Admitted(byte[]? data, string path) {
        if (data == null) throw new InvalidOperationException("Archive member '" + path + "' was not admitted.");
        ReaderNestedContent.MarkReservedPayload(data);
        return data;
    }
}
