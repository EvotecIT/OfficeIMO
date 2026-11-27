using OfficeIMO.Core.Internal;

namespace OfficeIMO.Archives;

public static partial class ArchiveExtractor {
    private static void ExtractTar(Stream source, string staging, ArchiveLimits limits, List<ExtractedArchiveFile> files,
        ref long total, CancellationToken token) {
        using TarSource tar = TarSource.Open(source, limits, token);
        foreach (var (entry, path) in tar.SelectFiles(limits, token, rejectLinks: true)) {
            token.ThrowIfCancellationRequested();
            using var payload = new SeekableReadWindowStream(tar.Stream, entry.DataOffset, entry.Length);
            WriteStream(staging, path, payload, entry.Length, limits, files, ref total, token);
        }
    }

    private static void ExtractGZip(Stream source, string package, string staging, ArchiveLimits limits,
        List<ExtractedArchiveFile> files, ref long total, CancellationToken token) {
        using var staged = new ArchiveSpillStream();
        source.Position = 0;
        GZipReadResult result = GZipContent.Decompress(source, staged, Math.Min(limits.Entry, limits.Expanded), limits,
            limits.Entry <= limits.Expanded ? nameof(ArchiveExtractionOptions.MaxEntryBytes) : nameof(ArchiveExtractionOptions.MaxExpandedBytes), token);
        limits.Operation.AdmitEntries(1);
        string name = !string.IsNullOrEmpty(result.FileName) ? result.FileName!
            : package.EndsWith(".gz", StringComparison.OrdinalIgnoreCase) && Path.GetFileName(package).Length > 3
                ? Path.GetFileNameWithoutExtension(package) : "content";
        string path = new ArchivePaths(limits).Add(name);
        if (!limits.Select(path)) return;
        limits.Operation.Select(staged.Length);
        staged.Position = 0;
        WriteStream(staging, path, staged, staged.Length, limits, files, ref total, token);
    }

    /// <summary>Extracts a single Zstandard, xz or lzip file. These formats store no file name, so the package name
    /// without its suffix is used, or "content".</summary>
    private static void ExtractCompressedFile(StreamCodec codec, Stream source, string package, string staging, ArchiveLimits limits,
        List<ExtractedArchiveFile> files, ref long total, CancellationToken token) {
        using var staged = new ArchiveSpillStream();
        source.Position = 0;
        StreamCodecs.Decompress(codec, source, staged, Math.Min(limits.Entry, limits.Expanded), limits,
            limits.Entry <= limits.Expanded ? nameof(ArchiveExtractionOptions.MaxEntryBytes) : nameof(ArchiveExtractionOptions.MaxExpandedBytes),
            selected: false, token);
        limits.Operation.AdmitEntries(1);
        string suffix = StreamCodecs.Suffix(codec);
        string file = Path.GetFileName(package);
        string name = file.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) && file.Length > suffix.Length
            ? file.Substring(0, file.Length - suffix.Length) : "content";
        string path = new ArchivePaths(limits).Add(name);
        if (!limits.Select(path)) return;
        limits.Operation.Select(staged.Length);
        staged.Position = 0;
        WriteStream(staging, path, staged, staged.Length, limits, files, ref total, token);
    }
}
