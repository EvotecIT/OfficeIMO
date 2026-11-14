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
}
