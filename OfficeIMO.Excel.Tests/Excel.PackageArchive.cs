using System.IO.Compression;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    /// <summary>Reads the package writer's output back with the platform ZIP reader at deflate block, region and history boundaries.</summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PackageArchive_EntriesRoundTripThroughThePlatformZipReader(bool seekable) {
        var random = new Random(20261110);
        byte[] noise = new byte[300_000];
        random.NextBytes(noise);
        byte[] rows = Rows(700_000);
        var parts = new List<(string Name, byte[] Content, CompressionLevel Level)> {
            ("empty.bin", Array.Empty<byte>(), CompressionLevel.Fastest),
            ("one.bin", new byte[] { 0x5a }, CompressionLevel.Fastest),
            ("stored-block.bin", rows.AsSpan(0, 65_535).ToArray(), CompressionLevel.Fastest),
            // Several 128 KiB regions, so matches reach back across the retained 32 KiB history.
            ("rows.bin", rows, CompressionLevel.Fastest),
            // Incompressible input takes the stored-block fallback.
            ("noise.bin", noise, CompressionLevel.Fastest),
            ("far-repeat.bin", noise.AsSpan(0, 32_768).ToArray().Concat(noise.AsSpan(0, 40_000).ToArray()).ToArray(), CompressionLevel.Fastest),
            ("ünïcode/naïve.xml", System.Text.Encoding.UTF8.GetBytes(string.Concat(Enumerable.Repeat("<c r=\"A1\"><v>1</v></c>", 5_000))), CompressionLevel.Fastest),
            ("stored.bin", noise.AsSpan(0, 5_000).ToArray(), CompressionLevel.NoCompression),
            ("optimal.bin", rows.AsSpan(0, 200_000).ToArray(), CompressionLevel.Optimal)
        };

        using var seekableOutput = new MemoryStream();
        var forwardOnly = new NonSeekableWriteStream();
        using (ExcelPackageArchive archive = ExcelPackageArchive.Create(seekable ? seekableOutput : forwardOnly)) {
            foreach ((string name, byte[] content, CompressionLevel level) in parts) {
                using Stream entry = archive.OpenEntry(name, level);
                // Uneven writes split blocks and regions at arbitrary offsets.
                for (int offset = 0, step = 1; offset < content.Length; offset += step, step = step * 3 % 100_003 + 1)
                    entry.Write(content, offset, Math.Min(step, content.Length - offset));
            }
        }

        using var zip = new ZipArchive(new MemoryStream(seekable ? seekableOutput.ToArray() : forwardOnly.ToArray()), ZipArchiveMode.Read);
        Assert.Equal(parts.Select(part => part.Name), zip.Entries.Select(entry => entry.FullName));
        foreach ((string name, byte[] content, _) in parts) {
            ZipArchiveEntry entry = zip.GetEntry(name)!;
            using var actual = new MemoryStream();
            using (Stream read = entry.Open()) read.CopyTo(actual);
            Assert.True(content.AsSpan().SequenceEqual(actual.ToArray()), name);
            Assert.Equal(content.Length, entry.Length);
#if NET
            Assert.Equal(ReferenceCrc32(content), entry.Crc32);
#endif
        }
        Assert.True(zip.GetEntry("rows.bin")!.CompressedLength < rows.Length / 2, "Repetitive rows were not compressed.");

        static byte[] Rows(int length) {
            var bytes = new byte[length];
            for (int index = 0, row = 0; index < length; row++) {
                byte[] record = System.Text.Encoding.ASCII.GetBytes("\u0000\u0019Name " + row + "\u0002" + (row * 1.25));
                int count = Math.Min(record.Length, length - index);
                Array.Copy(record, 0, bytes, index, count);
                index += count;
            }
            return bytes;
        }

        static uint ReferenceCrc32(byte[] bytes) {
            uint crc = 0xffffffff;
            foreach (byte value in bytes) {
                crc ^= value;
                for (int bit = 0; bit < 8; bit++) crc = (crc & 1) != 0 ? 0xEDB88320u ^ (crc >> 1) : crc >> 1;
            }
            return ~crc;
        }
    }
}
