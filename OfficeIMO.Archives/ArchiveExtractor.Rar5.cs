using OfficeIMO.Archives.Rar5;

namespace OfficeIMO.Archives;

public static partial class ArchiveExtractor {
    /// <summary>Verifies and extracts a complete RAR5 set supplied in volume order. Publishes only after all selected
    /// stored members pass integrity and path checks. Archive links and ownership are not applied on disk.</summary>
    public static ArchiveExtractionResult ExtractRar5Volumes(IReadOnlyList<string> volumePaths, string destinationDirectory,
        ArchiveExtractionOptions? options = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(volumePaths);
        ArgumentException.ThrowIfNullOrWhiteSpace(destinationDirectory);
        cancellationToken.ThrowIfCancellationRequested();
        int count = volumePaths.Count;
        if (count == 0) throw new ArgumentException("Supply at least one RAR5 volume.", nameof(volumePaths));
        ArchiveLimits.Check(count, Rar5Archive.MaxVolumeCount, "MaxVolumeCount");
        var limits = new ArchiveLimits(options ?? new ArchiveExtractionOptions());
        return Publish(ResolveDestination(destinationDirectory), limits, ArchiveFormat.Rar5, cancellationToken, ExtractInto);

        long ExtractInto(string staging, List<ExtractedArchiveFile> files) {
            var sources = new List<Stream>(count);
            long total = 0;
            try {
                for (int i = 0; i < count; i++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    string path = volumePaths[i];
                    ArgumentException.ThrowIfNullOrWhiteSpace(path);
                    var source = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
                    sources.Add(source);
                    ArchiveLimits.Add(ref total, source.Length, limits.Input, nameof(ArchiveExtractionOptions.MaxInputBytes));
                    limits.Operation.AdmitInput(source.Length);
                }
                return StageRar5(Rar5Archive.Open(sources, limits, cancellationToken), staging, limits, files, cancellationToken);
            } finally {
                foreach (Stream source in sources) source.Dispose();
            }
        }
    }

    private static long StageRar5(Rar5Archive archive, string staging, ArchiveLimits limits,
        List<ExtractedArchiveFile> files, CancellationToken token) {
        long staged = 0;
        var buffer = new byte[StagingCopyBufferSize];
        archive.Read((path, data) => {
            using var payload = new MemoryStream(data, writable: false);
            WriteStream(staging, path, payload, data.LongLength, limits, files, ref staged, token, buffer);
        }, token, path => {
            string target = Path.GetFullPath(Path.Combine(staging, path.Replace('/', Path.DirectorySeparatorChar)));
            ArchivePaths.RejectLinks(target, staging);
            Directory.CreateDirectory(target);
            ArchivePaths.RejectLinks(target, staging);
        });
        return staged;
    }
}
