using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace OfficeIMO.Archives;

public static partial class ArchiveExtractor {
    private static void ExtractZip(Stream source, string staging, ArchiveLimits limits, List<ExtractedArchiveFile> files,
        ref long total, CancellationToken token) {
        long written = total;
        ZipArchiveContent.Read(source, limits, (path, payload, length) =>
            WriteStream(staging, path, payload, length, limits, files, ref written, token), token);
        total = written;
    }
}
