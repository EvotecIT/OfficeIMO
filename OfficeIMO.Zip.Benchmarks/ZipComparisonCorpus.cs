using System.IO.Compression;

namespace OfficeIMO.Zip.Benchmarks;

internal static class ZipComparisonCorpus {
    internal static readonly ZipBenchmarkScale[] Scales = {
        new("Small", 24, 4, 96),
        new("Normal", 512, 32, 160),
        new("Large", 4000, 64, 224)
    };

    internal static IEnumerable<string> ScaleNames => Scales.Select(scale => scale.Name);

    internal static ZipBenchmarkScale Get(string name) => Scales.FirstOrDefault(
        scale => string.Equals(scale.Name, name, StringComparison.OrdinalIgnoreCase))
        ?? throw new ArgumentException("Unknown ZIP benchmark scale: " + name, nameof(name));

    internal static byte[] CreateArchive(ZipBenchmarkScale scale) {
        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            for (int directory = scale.DirectoryCount - 1; directory >= 0; directory--) {
                ZipArchiveEntry entry = archive.CreateEntry($"group-{directory:D3}/");
                entry.LastWriteTime = FixedTimestamp;
            }

            for (int index = scale.FileCount - 1; index >= 0; index--) {
                int group = index % scale.DirectoryCount;
                int subgroup = (index / scale.DirectoryCount) % 8;
                ZipArchiveEntry entry = archive.CreateEntry(
                    $"group-{group:D3}/part-{subgroup:D2}/entry-{index:D5}.bin",
                    CompressionLevel.Fastest);
                entry.LastWriteTime = FixedTimestamp;
                using Stream destination = entry.Open();
                byte[] payload = CreatePayload(scale.PayloadBytes, index);
                destination.Write(payload, 0, payload.Length);
            }
        }
        return output.ToArray();
    }

    private static byte[] CreatePayload(int length, int seed) {
        var payload = new byte[length];
        uint state = unchecked((uint)(seed + 1) * 2654435761U);
        for (int index = 0; index < payload.Length; index++) {
            state = unchecked(state * 1664525U + 1013904223U);
            payload[index] = (byte)(state >> 24);
        }
        return payload;
    }

    private static readonly DateTimeOffset FixedTimestamp = new(2024, 1, 2, 3, 4, 6, TimeSpan.Zero);
}

internal sealed record ZipBenchmarkScale(string Name, int FileCount, int DirectoryCount, int PayloadBytes);
