using System.IO;
using System.Threading;

namespace OfficeIMO.Core.Internal;

internal static partial class OfficeArchiveSafety {
    // Retain the published workflow call while sharing the bounded, position-preserving scanner.
    internal static ZipCentralDirectoryScanResult ScanZipCentralDirectory(
        Stream source, long archiveLength, int entryLimit) =>
        ScanZipCentralDirectory(source, archiveLength, entryLimit, CancellationToken.None);
}
