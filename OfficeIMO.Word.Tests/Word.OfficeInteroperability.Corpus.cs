using OfficeIMO.TestAssets;
using OfficeIMO.Word;
using OfficeIMO.Word.LegacyDoc;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        [Trait("Category", "OfficeInteroperability")]
        public void OfficeInteroperability_CorpusManifest_TracksAndLoadsWordArtifacts() {
            OfficeInteroperabilityCorpusManifest manifest = OfficeInteroperabilityCorpusManifestLoader.Load();
            IReadOnlyList<string> errors = OfficeInteroperabilityCorpusManifestLoader.Validate(manifest);
            Assert.True(errors.Count == 0, string.Join(Environment.NewLine, errors));

            OfficeInteroperabilityCorpusCollection collection = Assert.Single(
                manifest.Collections,
                item => item.Format == "doc");
            Assert.Equal("compatibility", collection.Role);
            Assert.Equal("Microsoft Word COM", collection.Producer);

            foreach (OfficeInteroperabilityCorpusArtifact artifact in collection.Artifacts) {
                string path = OfficeInteroperabilityCorpusManifestLoader.ResolveArtifactPath(collection, artifact);
                using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(path);
                result.EnsureNoImportErrors();
                Assert.True(result.HasDocument, artifact.File);
                Assert.Equal(WordFileFormat.Doc, result.Document.SourceFormat);
                Assert.NotEmpty(result.Document.Paragraphs);
            }
        }
    }
}
