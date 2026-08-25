using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.Zip.Benchmarks;

internal static class ZipComparisonWorkflows {
    internal static ZipTraversalResult TraverseOffice(byte[] input) {
        using var stream = new MemoryStream(input, writable: false);
        return ZipTraversal.Traverse(stream);
    }

    internal static IReadOnlyList<ZipProjectionDescriptor> TraversePlatform(byte[] input) {
        using var stream = new MemoryStream(input, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
        var descriptors = new List<ZipProjectionDescriptor>(archive.Entries.Count);
        foreach (ZipArchiveEntry entry in archive.Entries.OrderBy(item => item.FullName, StringComparer.Ordinal)) {
            string fullName = entry.FullName;
            bool isDirectory = fullName.EndsWith("/", StringComparison.Ordinal);
            if (isDirectory) continue;
            descriptors.Add(new ZipProjectionDescriptor(
                fullName,
                entry.Name,
                false,
                ComputeDepth(fullName),
                entry.Length,
                entry.LastWriteTime.UtcDateTime));
        }
        return descriptors;
    }

    internal static ZipComparisonObservation Observe(IReadOnlyList<ZipEntryDescriptor> entries) => Observe(
        entries.Select(entry => new ZipProjectionDescriptor(
            entry.FullName,
            entry.Name,
            entry.IsDirectory,
            entry.Depth,
            entry.UncompressedLength,
            entry.LastWriteUtc)));

    internal static ZipComparisonObservation Observe(IReadOnlyList<ZipProjectionDescriptor> entries) => Observe(
        entries.AsEnumerable());

    private static ZipComparisonObservation Observe(IEnumerable<ZipProjectionDescriptor> entries) {
        int count = 0;
        long totalBytes = 0;
        using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        foreach (ZipProjectionDescriptor entry in entries) {
            count++;
            totalBytes += entry.UncompressedLength;
            byte[] line = Encoding.UTF8.GetBytes(
                $"{entry.FullName}\0{entry.Name}\0{entry.IsDirectory}\0{entry.Depth}\0" +
                $"{entry.UncompressedLength}\0{entry.LastWriteUtc.Ticks}\n");
            hash.AppendData(line);
        }
        return new ZipComparisonObservation(count, totalBytes, Convert.ToHexString(hash.GetHashAndReset()));
    }

    private static int ComputeDepth(string fullName) {
        int depth = 1;
        for (int index = 0; index < fullName.Length; index++) {
            if (fullName[index] == '/') depth++;
        }
        return depth;
    }
}

public sealed record ZipProjectionDescriptor(
    string FullName,
    string Name,
    bool IsDirectory,
    int Depth,
    long UncompressedLength,
    DateTime LastWriteUtc);

internal sealed record ZipComparisonObservation(
    int EntryCount,
    long TotalUncompressedBytes,
    string StructuralFingerprint);
