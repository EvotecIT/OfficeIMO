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
            var chunks = reader.Read(input, "input.zip", new() { ComputeHashes = false }).ToArray();
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
            var chunks = reader.Read(input, "input.zip", new() { ComputeHashes = false }).ToArray();
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
            var error = Assert.Throws<ReaderResourceLimitException>(() => reader.Read(input, "input.zip").ToArray());
            Assert.IsType<OfficeInspectionResourceLimitException>(error.InnerException);
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
