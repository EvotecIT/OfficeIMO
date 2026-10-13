namespace OfficeIMO.Archives.Tests {

    public sealed class ArchivePortablePathTests {
        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void Utf8ByteOversizedSegmentsAreRejectedBeforeCabinetSelection(bool supplementary) {
            string name = supplementary ? string.Concat(Enumerable.Repeat("\U0001f600", 64)) : new string('é', 128);
            Assert.Equal(256, Encoding.UTF8.GetByteCount(name));
            var paths = new ArchivePaths(new ArchiveLimits(new ArchiveExtractionOptions()));
            Assert.Throws<InvalidDataException>(() => paths.Add(name));
            Assert.Throws<InvalidDataException>(() => paths.Add(name + "/template.admx"));
            using var fixture = new ArchiveFixtures();
            int selectionCalls = 0;
            Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(
                fixture.Package(ArchiveFixtures.Utf8Cabinet(name)), fixture.Destination,
                new ArchiveExtractionOptions { IncludeFile = _ => { selectionCalls++; return false; } }));
            Assert.Equal(0, selectionCalls);
            Assert.False(Directory.Exists(fixture.Destination));
            Assert.Empty(Directory.GetDirectories(fixture.Root, ".officeimo-archive-*"));
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void Utf8SegmentsAtTheByteBoundaryPreserveAuthoredSpelling(bool supplementary) {
            string name = supplementary ? string.Concat(Enumerable.Repeat("\U0001f600", 63)) + "abc" : new string('é', 127) + "x";
            Assert.Equal(255, Encoding.UTF8.GetByteCount(name));
            using var fixture = new ArchiveFixtures();
            var result = ArchiveExtractor.Extract(fixture.Package(ArchiveFixtures.Utf8Cabinet(name)), fixture.Destination);
            Assert.Equal(name, Assert.Single(result.Files).RelativePath);
            Assert.Equal("A", File.ReadAllText(Path.Combine(fixture.Destination, name)));
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void MsiFilenameAndDirectorySegmentsShareTheUtf8BytePreflight(bool directory) {
            string segment = new('é', 128);
            using var fixture = new ArchiveFixtures();
            var file = directory ? new MsiFixtureFile("Template.adml", parent: segment) : new MsiFixtureFile(segment);
            int selectionCalls = 0;
            Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(
                fixture.Package(ArchiveFixtures.MsiAliasFiles(false, file)), fixture.Destination,
                new ArchiveExtractionOptions { IncludeFile = _ => { selectionCalls++; return false; } }));
            Assert.Equal(0, selectionCalls);
            Assert.False(Directory.Exists(fixture.Destination));
            Assert.Empty(Directory.GetDirectories(fixture.Root, ".officeimo-archive-*"));
        }
    }
}
