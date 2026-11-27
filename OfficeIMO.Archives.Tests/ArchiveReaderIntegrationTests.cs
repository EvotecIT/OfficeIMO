using System.IO.Compression;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Zip;

namespace OfficeIMO.Archives.Tests {

    public sealed class ArchiveReaderIntegrationTests {
        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void CorruptStoredMembersNeverReachDocumentHandlers(bool nested) {
            byte[] bytes = ArchiveZipInspectionTests.Zip(("test.txt", "original"));
            bytes[ArchiveZipInspectionTests.Payload(bytes)] ^= 1;
            if (nested) bytes = Nest(bytes);
            var reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers().AddZipHandler().Build();
            using var input = new MemoryStream(bytes);
            var chunks = reader.Enumerate(input, "input.zip", new() { ComputeHashes = false }).ToArray();
            Assert.NotEmpty(chunks);
            Assert.All(chunks, c => Assert.NotEmpty(c.Warnings ?? Array.Empty<string>()));
            Assert.DoesNotContain(chunks, c => c.Text.Contains("original") || c.Text.Contains("nriginal"));
        }

        [Fact]
        public void DeflateSizeLiesCannotBeAcceptedAsTruncatedDocumentContent() {
            using var output = new MemoryStream();
            using (var zip = new ZipArchive(output, ZipArchiveMode.Create, true)) {
                using var writer = new StreamWriter(zip.CreateEntry("test.txt", CompressionLevel.Optimal).Open(), new UTF8Encoding(false));
                writer.Write("LONGER CONTENT");
            }
            byte[] bytes = output.ToArray();
            BitConverter.GetBytes(2u).CopyTo(bytes, 22);
            BitConverter.GetBytes(2u).CopyTo(bytes, ArchiveZipInspectionTests.Central(bytes) + 24);
            var reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers().AddZipHandler().Build();
            using var input = new MemoryStream(bytes);
            var chunks = reader.Enumerate(input, "input.zip", new() { ComputeHashes = false }).ToArray();
            Assert.NotEmpty(chunks);
            Assert.All(chunks, c => Assert.NotEmpty(c.Warnings ?? Array.Empty<string>()));
            Assert.DoesNotContain(chunks, c => c.Text == "LO");
        }

        [Fact]
        public void NestedArchiveBudgetExhaustionIsTerminalInsteadOfARecoverableWarning() {
            var budget = new ArchiveOperationBudget(new() { MaxEntryCount = 1 });
            var reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers()
                .AddZipHandler(new() { OperationBudget = budget }).Build();
            using var input = new MemoryStream(Nest(ArchiveZipInspectionTests.Zip(("test.txt", "hello"))));
            var error = Assert.Throws<ReaderResourceLimitException>(() => reader.Enumerate(input, "input.zip").ToArray());
            Assert.IsType<OfficeInspectionResourceLimitException>(error.InnerException);
        }

        [Theory]
        [InlineData("input.tar", false)]
        [InlineData("input.tgz", true)]
        [InlineData("input.tar.gz", true)]
        public void TarMembersReachDocumentHandlersWithVirtualPaths(string name, bool gzip) {
            using var archive = new MemoryStream();
            TarArchiveWriter.Create(archive, new[] {
                ArchiveWriteEntry.File("docs/readme.txt", Encoding.UTF8.GetBytes("hello from tar")),
                ArchiveWriteEntry.File("docs/skip.bin", new byte[] { 1, 2, 3 })
            }, new TarArchiveWriterOptions { GZip = gzip });
            var reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers().AddZipHandler().Build();
            archive.Position = 0;
            var chunks = reader.Enumerate(archive, name, new() { ComputeHashes = false }).ToArray();
            var chunk = Assert.Single(chunks);
            Assert.Contains("hello from tar", chunk.Text);
            Assert.Equal(name + "::docs/readme.txt", chunk.Location.Path);
        }

        [Theory]
        [InlineData("archive.tar.zst")]
        [InlineData("archive.tar.xz")]
        [InlineData("archive.tar.lz")]
        public void CompressedTarMembersReachDocumentHandlers(string name) {
            byte[] bytes = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "IndependentTarWrappers", name));
            var reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers().AddZipHandler().Build();
            var chunks = reader.Enumerate(new MemoryStream(bytes), name, new() { ComputeHashes = false }).ToArray();
            Assert.Contains(chunks, chunk => chunk.Location.Path == name + "::docs/text.txt" && chunk.Text.Contains("Line 0 of the harbour log"));
        }

        [Theory]
        [InlineData("IndependentZstandard", "text-level1.zst", ".zst")]
        [InlineData("IndependentXzLzip", "text-preset0.xz", ".xz")]
        [InlineData("IndependentXzLzip", "text-level0.lz", ".lz")]
        public void SingleCompressedFilesAreReadAsTheirInnerDocument(string folder, string fixture, string suffix) {
            byte[] bytes = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", folder, fixture));
            var reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers().AddZipHandler().Build();
            var chunks = reader.Enumerate(new MemoryStream(bytes), "log.txt" + suffix, new() { ComputeHashes = false }).ToArray();
            Assert.NotEmpty(chunks);
            Assert.All(chunks, chunk => Assert.StartsWith("log.txt" + suffix + "::log.txt", chunk.Location.Path));
            Assert.Contains("Line 0 of the harbour log", chunks[0].Text);
        }

        [Fact]
        public void SingleFileGZipIsReadAsItsInnerDocument() {
            using var archive = new MemoryStream();
            GZipArchive.Compress(new MemoryStream(Encoding.UTF8.GetBytes("compressed notes")), archive);
            archive.Position = 0;
            var reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers().AddZipHandler().Build();
            var chunk = Assert.Single(reader.Enumerate(archive, "notes.txt.gz", new() { ComputeHashes = false }).ToArray());
            Assert.Contains("compressed notes", chunk.Text);
            Assert.Equal("notes.txt.gz::notes.txt", chunk.Location.Path);
        }

        [Fact]
        public void SingleFileGZipHonoursTheTotalExpansionLimitAndTheSharedSelectionBudget() {
            using var archive = new MemoryStream();
            GZipArchive.Compress(new MemoryStream(Encoding.UTF8.GetBytes(new string('a', 4096))), archive);
            var limited = new OfficeDocumentReaderBuilder().AddPlainTextHandlers().AddZipHandler(new() { MaxTotalUncompressedBytes = 1024 }).Build();
            archive.Position = 0;
            Assert.Throws<ReaderResourceLimitException>(() => limited.Enumerate(archive, "notes.txt.gz", new() { ComputeHashes = false }).ToArray());
            var budget = new ArchiveOperationBudget(new() { MaxSelectedBytes = 4 });
            var budgeted = new OfficeDocumentReaderBuilder().AddPlainTextHandlers().AddZipHandler(new() { OperationBudget = budget }).Build();
            archive.Position = 0;
            Assert.Throws<ReaderResourceLimitException>(() => budgeted.Enumerate(archive, "notes.txt.gz", new() { ComputeHashes = false }).ToArray());
        }

        [Fact]
        public void NestedInputLimitCoversEveryTarMember() {
            using var archive = new MemoryStream();
            TarArchiveWriter.Create(archive, new[] {
                ArchiveWriteEntry.File("a.txt", Encoding.UTF8.GetBytes(new string('a', 3000))),
                ArchiveWriteEntry.File("b.txt", Encoding.UTF8.GetBytes(new string('b', 3000)))
            });
            var reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers().AddZipHandler().Build();
            ReaderOptions Limit(long bytes) => new() { ComputeHashes = false, ResourceLimits = new ReaderResourceLimits { MaxNestedInputBytes = bytes } };
            archive.Position = 0;
            Assert.Throws<ReaderResourceLimitException>(() => reader.Enumerate(archive, "input.tar", Limit(4000)).ToArray());
            archive.Position = 0;
            Assert.Equal(2, reader.Enumerate(archive, "input.tar", Limit(6000)).Count());
        }

        private static byte[] Nest(byte[] bytes) {
            using var output = new MemoryStream();
            using (var zip = new ZipArchive(output, ZipArchiveMode.Create, true)) {
                using var member = zip.CreateEntry("nested.zip", CompressionLevel.NoCompression).Open(); member.Write(bytes, 0, bytes.Length);
            }
            return output.ToArray();
        }
    }
}
