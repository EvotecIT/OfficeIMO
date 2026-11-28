using OfficeIMO.Archives.SevenZip;

namespace OfficeIMO.Archives;

public static partial class ArchiveExtractor {
    private static void ExtractSevenZip(Stream source, string staging, ArchiveLimits limits, List<ExtractedArchiveFile> files,
        ref long total, CancellationToken token) {
        long stagedBytes = total;
        var transferBuffer = new byte[StagingCopyBufferSize];
        SevenZipArchive.Open(source, limits, token).Read((path, buffer, offset, count) => {
            using var payload = new MemoryStream(buffer, offset, count, writable: false);
            WriteStream(staging, path, payload, count, limits, files, ref stagedBytes, token, transferBuffer);
        }, token);
        total = stagedBytes;
    }
}
