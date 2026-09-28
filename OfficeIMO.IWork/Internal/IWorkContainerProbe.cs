using System.IO.Compression;
using System.Threading;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.IWork.Internal;

/// <summary>Checks iWork ZIP structure without expanding package entries or parsing IWA records.</summary>
internal static class IWorkContainerProbe {
    internal static bool HasModernIndex(Stream stream, long maximumPackageBytes,
        int maximumEntries, CancellationToken cancellationToken) {
        if (!stream.CanRead || !stream.CanSeek) return false;
        long start = stream.Position;
        try {
            long packageLength = stream.Length - start;
            if (packageLength > maximumPackageBytes) return false;
            using var window = new SeekableReadWindowStream(stream, start, packageLength);
            OfficeArchiveSafety.ZipCentralDirectoryScanResult directory =
                OfficeArchiveSafety.ScanZipCentralDirectory(window,
                    packageLength, maximumEntries, cancellationToken);
            if (!directory.IsValid || directory.LimitExceeded) return false;
            using var archive = new ZipArchive(window, ZipArchiveMode.Read, leaveOpen: true);
            foreach (ZipArchiveEntry entry in archive.Entries) {
                cancellationToken.ThrowIfCancellationRequested();
                string path = entry.FullName.Replace('\\', '/');
                if (string.Equals(path, "Index.zip", StringComparison.OrdinalIgnoreCase)) {
                    return entry.Length > 0;
                }
                if (string.Equals(path, "Index/Document.iwa", StringComparison.OrdinalIgnoreCase)) {
                    return entry.Length > 0;
                }
            }
            return false;
        } catch (InvalidDataException) {
            return false;
        } catch (IOException) {
            return false;
        } finally {
            stream.Position = start;
        }
    }
}
