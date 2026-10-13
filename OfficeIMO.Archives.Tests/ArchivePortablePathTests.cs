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

        [Theory]
        [InlineData("Windows Server 2022 August 2021 Update ", "Windows Server 2022 August 2021 Update")]
        [InlineData("Templates.", "Templates")]
        [InlineData("Templates. .", "Templates")]
        public void MsiDirectorySegmentsUseTheirInstalledSpelling(string authored, string installed) {
            using var fixture = new ArchiveFixtures();
            var result = ArchiveExtractor.Extract(fixture.Package(ArchiveFixtures.MsiAliasFiles(false,
                new MsiFixtureFile("Template.admx", "A", parent: authored + "/PolicyDefinitions"))), fixture.Destination);
            string relative = installed + "/PolicyDefinitions/Template.admx";
            Assert.Equal(relative, Assert.Single(result.Files).RelativePath);
            Assert.Equal("A", File.ReadAllText(Path.Combine(fixture.Destination, relative.Replace('/', Path.DirectorySeparatorChar))));
        }

        [Theory]
        [InlineData("..")]
        [InlineData("...")]
        [InlineData("   ")]
        [InlineData(". .")]
        [InlineData("CON ")]
        [InlineData("NUL.")]
        [InlineData("a*b ")]
        public void MsiDirectorySegmentsStillRejectTraversalAndNonportableNamesAfterNormalization(string authored) {
            using var fixture = new ArchiveFixtures();
            int selectionCalls = 0;
            Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(
                fixture.Package(ArchiveFixtures.MsiAliasFiles(false, new MsiFixtureFile("Template.admx", parent: authored + "/PolicyDefinitions"))),
                fixture.Destination, new ArchiveExtractionOptions { IncludeFile = _ => { selectionCalls++; return true; } }));
            Assert.Equal(0, selectionCalls);
            Assert.False(Directory.Exists(fixture.Destination));
            Assert.Empty(Directory.GetDirectories(fixture.Root, ".officeimo-archive-*"));
        }

        [Fact]
        public void MsiDirectoryNormalizationAppliesToTheLongTargetName() {
            using var fixture = new ArchiveFixtures();
            var result = ArchiveExtractor.Extract(fixture.Package(ArchiveFixtures.MsiAliasFiles(false,
                new MsiFixtureFile("Template.admx", "A", parent: "TEMPLA~1|Templates :SOURCE~1|Source "))), fixture.Destination);
            Assert.Equal("Templates/Template.admx", Assert.Single(result.Files).RelativePath);
        }

        [Fact]
        public void MsiDirectoriesThatNormalizeToTheSameSpellingMergeLikeWindowsInstaller() {
            using var fixture = new ArchiveFixtures();
            var result = ArchiveExtractor.Extract(fixture.Package(ArchiveFixtures.MsiAliasFiles(false,
                new MsiFixtureFile("First.admx", "A", parent: "Templates "),
                new MsiFixtureFile("Second.admx", "B", parent: "Templates"))), fixture.Destination);
            Assert.Equal(new[] { "Templates/First.admx", "Templates/Second.admx" },
                result.Files.Select(file => file.RelativePath).OrderBy(path => path, StringComparer.Ordinal));
        }

        [Fact]
        public void MsiDirectoriesThatNormalizeToTheSameSpellingRejectConflictingFiles() {
            using var fixture = new ArchiveFixtures();
            Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(fixture.Package(ArchiveFixtures.MsiAliasFiles(false,
                new MsiFixtureFile("Template.admx", "A", parent: "Templates "),
                new MsiFixtureFile("Template.admx", "B", parent: "Templates"))), fixture.Destination));
            Assert.False(Directory.Exists(fixture.Destination));
        }

        [Fact]
        public void MsiDirectoriesThatNormalizeToCaseVariantsStillCollide() {
            using var fixture = new ArchiveFixtures();
            Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(fixture.Package(ArchiveFixtures.MsiAliasFiles(false,
                new MsiFixtureFile("First.admx", "A", parent: "Templates "),
                new MsiFixtureFile("Second.admx", "B", parent: "templates"))), fixture.Destination));
            Assert.False(Directory.Exists(fixture.Destination));
        }

        [Fact]
        public void MsiDirectoryThatNormalizesToASiblingFileNameStillCollides() {
            using var fixture = new ArchiveFixtures();
            Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(fixture.Package(ArchiveFixtures.MsiAliasFiles(false,
                new MsiFixtureFile("Template.admx", "A", parent: "Root/Templates "),
                new MsiFixtureFile("Templates", "B", parent: "Root"))), fixture.Destination));
            Assert.False(Directory.Exists(fixture.Destination));
        }
    }
}
