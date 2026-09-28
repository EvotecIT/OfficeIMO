using System.IO.Compression;

namespace OfficeIMO.IWork.Internal;

/// <summary>Checks iWork ZIP structure without expanding package entries or parsing IWA records.</summary>
internal static class IWorkContainerProbe {
    internal static bool HasModernIndex(Stream stream, long maximumPackageBytes) {
        if (!stream.CanRead || !stream.CanSeek) return false;
        long start = stream.Position;
        try {
            if (stream.Length - start > maximumPackageBytes) return false;
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
            foreach (ZipArchiveEntry entry in archive.Entries) {
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
