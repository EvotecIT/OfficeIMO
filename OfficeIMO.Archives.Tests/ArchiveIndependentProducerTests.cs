namespace OfficeIMO.Archives.Tests {

    public sealed class ArchiveIndependentProducerTests {
        [Theory]
        [InlineData("7zip-utf8.zip", 0)]
        [InlineData("7zip-unflagged-utf8.zip", 65001)]
        [InlineData("python-cp1250.zip", 1250)]
        [InlineData("python-streaming.zip", 0)]
        [InlineData("python-streaming-zip64.zip", 0)]
        public void IndependentZipProducersRetainNamesAndCompletePayloadsAcrossPublicRoutes(string fixture, int codePage) {
            string path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "IndependentZip", fixture);
            byte[] bytes = File.ReadAllBytes(path);
            int? encoding = codePage == 0 ? null : codePage;
            if (codePage != 0) {
                Assert.Throws<NotSupportedException>(() => ArchiveReader.ReadZip(bytes));
                Assert.Throws<NotSupportedException>(() => ArchiveReader.TraverseZip(path));
                Assert.Throws<NotSupportedException>(() => ArchiveReader.TraverseZip(new MemoryStream(bytes)));
            }
            Assert.Equal("zażółć-gęślą-jaźń.txt", Assert.Single(ArchiveReader.TraverseZip(path,
                new() { ZipLegacyCodePage = encoding }).Entries).FullName);
            using var inspection = ArchiveInspector.Open(path, new() { ZipLegacyCodePage = encoding });
            var entry = Assert.Single(inspection.Entries);
            Assert.Equal("zażółć-gęślą-jaźń.txt", entry.RawName);
            using var batch = inspection.ReadMembers(new[] { entry.Ordinal }, 1024);
            byte[] expected = Encoding.UTF8.GetBytes("Polish filename and payload: Zażółć gęślą jaźń.\n");
            using var actual = new MemoryStream(); batch.Members[0].Content.CopyTo(actual);
            Assert.Equal(expected, actual.ToArray());
            Assert.Equal(expected, Assert.Single(ArchiveReader.ReadZip(bytes, new() { ZipLegacyCodePage = encoding })).Data);
            using var output = new ArchiveFixtures();
            var result = ArchiveExtractor.Extract(path, output.Destination, new() { ZipLegacyCodePage = encoding });
            Assert.Equal(expected, File.ReadAllBytes(Path.Combine(output.Destination, Assert.Single(result.Files).RelativePath)));
        }
    }
}
