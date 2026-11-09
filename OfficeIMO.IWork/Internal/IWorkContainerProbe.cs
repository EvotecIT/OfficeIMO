using System.IO.Compression;
using System.Xml;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.IWork.Internal;

/// <summary>Checks iWork ZIP structure using bounded modern or legacy index inspection.</summary>
internal static class IWorkContainerProbe {
    internal static bool HasSupportedIndex(Stream stream, IWorkDocumentKind expectedKind, long maximumPackageBytes,
        int maximumEntries, CancellationToken cancellationToken,
        long maximumIndexBytes = 128L * 1024 * 1024) {
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
            ZipArchiveEntry? nestedIndex = null;
            ZipArchiveEntry? legacyIndex = null;
            bool conflictingLegacyIndexes = false;
            foreach (ZipArchiveEntry entry in archive.Entries) {
                cancellationToken.ThrowIfCancellationRequested();
                string path = entry.FullName.Replace('\\', '/');
                if (string.Equals(path, "Index.zip", StringComparison.OrdinalIgnoreCase)) {
                    nestedIndex ??= entry;
                }
                if (string.Equals(path, "Index/Document.iwa", StringComparison.OrdinalIgnoreCase)) {
                    return entry.Length > 0;
                }
                if (IWorkLegacyXmlSource.IsIndexPath(path)) {
                    conflictingLegacyIndexes |= legacyIndex != null;
                    legacyIndex ??= entry;
                }
            }
            if (nestedIndex != null && HasNestedDocumentIndex(nestedIndex,
                    maximumIndexBytes, maximumEntries, cancellationToken)) return true;
            if (conflictingLegacyIndexes || legacyIndex == null
                || legacyIndex.Length <= 0 || legacyIndex.Length > maximumIndexBytes) return false;
            using Stream index = legacyIndex.Open();
            return IWorkLegacyXmlSource.Probe(index, legacyIndex.FullName, expectedKind, maximumIndexBytes,
                cancellationToken);
        } catch (InvalidDataException) {
            return false;
        } catch (XmlException) {
            return false;
        } catch (IOException) {
            return false;
        } catch (NotSupportedException) {
            return false;
        } finally {
            stream.Position = start;
        }
    }

    private static bool HasNestedDocumentIndex(ZipArchiveEntry entry, long maximumBytes,
        int maximumEntries, CancellationToken cancellationToken) {
        if (entry.Length <= 0 || entry.Length > maximumBytes) return false;
        using var content = new MemoryStream();
        using (Stream input = entry.Open()) {
            var buffer = new byte[81920];
            while (true) {
                cancellationToken.ThrowIfCancellationRequested();
                int read = input.Read(buffer, 0, buffer.Length);
                if (read == 0) break;
                if (content.Length > maximumBytes - read) return false;
                content.Write(buffer, 0, read);
            }
        }
        if (content.Length != entry.Length) return false;
        content.Position = 0;
        OfficeArchiveSafety.ZipCentralDirectoryScanResult directory =
            OfficeArchiveSafety.ScanZipCentralDirectory(content, content.Length,
                maximumEntries, cancellationToken);
        if (!directory.IsValid || directory.LimitExceeded) return false;
        content.Position = 0;
        using var nested = new ZipArchive(content, ZipArchiveMode.Read, leaveOpen: true);
        foreach (ZipArchiveEntry candidate in nested.Entries) {
            cancellationToken.ThrowIfCancellationRequested();
            string path = candidate.FullName.Replace('\\', '/');
            if ((string.Equals(path, "Document.iwa", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(path, "Index/Document.iwa", StringComparison.OrdinalIgnoreCase))
                && candidate.Length > 0) return true;
        }
        return false;
    }
}
