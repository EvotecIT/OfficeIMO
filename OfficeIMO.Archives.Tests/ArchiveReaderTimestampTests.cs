using System.IO.Compression;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Zip;

namespace OfficeIMO.Archives.Tests {

    public sealed class ArchiveReaderTimestampTests {
        [Theory]
        [InlineData(false, false)]
        [InlineData(false, true)]
        [InlineData(true, false)]
        [InlineData(true, true)]
        public void MemberMetadataMatchesPlatformZipTimeForContentAndReadWarnings(bool nested, bool corrupt) {
            byte[] bytes = CreateZip("test.txt", Encoding.UTF8.GetBytes("content"), out DateTime expected);
            if (corrupt) bytes[ArchiveZipInspectionTests.Payload(bytes)] ^= 1;
            if (nested) bytes = CreateZip("nested.zip", bytes, out _);
            var reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers().AddZipHandler().Build();
            using var input = new MemoryStream(bytes);
            var chunk = Assert.Single(reader.Read(input, "input.zip", new() { ComputeHashes = false }));
            Assert.Equal(expected, chunk.SourceLastWriteUtc);
            Assert.EndsWith("::test.txt", chunk.Location.Path);
            Assert.Equal(corrupt, chunk.Warnings?.Any() ?? false);
        }

        [Fact]
        public void DisabledNestedArchiveWarningUsesTheContainingEntryPlatformTime() {
            byte[] bytes = CreateZip("nested.zip", CreateZip("test.txt", Encoding.UTF8.GetBytes("content"), out _), out DateTime expected);
            var reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers()
                .AddZipHandler(readerZipOptions: new() { ReadNestedZipEntries = false }).Build();
            using var input = new MemoryStream(bytes);
            var chunk = Assert.Single(reader.Read(input, "input.zip", new() { ComputeHashes = false }));
            Assert.Equal(expected, chunk.SourceLastWriteUtc);
            Assert.NotEmpty(chunk.Warnings!);
            Assert.EndsWith("::nested.zip", chunk.Location.Path);
        }

        private static byte[] CreateZip(string name, byte[] content, out DateTime platformTime) {
            using var output = new MemoryStream();
            using (var zip = new ZipArchive(output, ZipArchiveMode.Create, true)) {
                var entry = zip.CreateEntry(name, CompressionLevel.NoCompression);
                entry.LastWriteTime = new DateTimeOffset(name.EndsWith(".zip") ? 2032 : 2024, 1, 2, 3, 4, 6, TimeSpan.Zero);
                using var member = entry.Open(); member.Write(content, 0, content.Length);
            }
            output.Position = 0;
            using var oracle = new ZipArchive(output, ZipArchiveMode.Read, true);
            platformTime = oracle.Entries[0].LastWriteTime.UtcDateTime;
            return output.ToArray();
        }
    }
}
