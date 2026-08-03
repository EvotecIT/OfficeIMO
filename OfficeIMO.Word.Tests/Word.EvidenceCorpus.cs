using OfficeIMO.TestAssets;
using OfficeIMO.Word;
using OfficeIMO.Word.LegacyDoc;
using System.Text.Json;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        [Trait("Category", "OfficeInteroperability")]
        public void WordEvidenceCorpus_TracksRequiredFamiliesProvenanceHashesAndLossPolicies() {
            WordEvidenceCorpusManifest manifest = WordEvidenceCorpusManifestLoader.Load();

            IReadOnlyList<string> errors = WordEvidenceCorpusManifestLoader.Validate(manifest);

            Assert.True(errors.Count == 0, string.Join(Environment.NewLine, errors));
            Assert.Contains(manifest.Artifacts, artifact => artifact.Id == "rendering-anchored-group-word-oracle" && artifact.Path == null);
            Assert.Contains(manifest.Artifacts, artifact => artifact.Families.Contains("legacy-doc") && artifact.LossPolicy == "guarded");
            Assert.All(manifest.Artifacts.Where(artifact => artifact.Path?.EndsWith(".docx", StringComparison.OrdinalIgnoreCase) == true ||
                                                           artifact.Path?.EndsWith(".doc", StringComparison.OrdinalIgnoreCase) == true),
                artifact => Assert.Equal("raw", artifact.HashMode));
            Assert.All(manifest.Artifacts.Where(artifact => artifact.Path?.StartsWith("Word/EvidenceCorpus/", StringComparison.OrdinalIgnoreCase) == true),
                artifact => Assert.Equal("canonical-text", artifact.HashMode));
            Assert.DoesNotContain(manifest.Artifacts, artifact =>
                artifact.Contract.IndexOf("native DOC authoring is supported", StringComparison.OrdinalIgnoreCase) >= 0);
        }

        [Fact]
        [Trait("Category", "OfficeInteroperability")]
        public void WordEvidenceCorpus_LoadsProducerArtifactsThroughTheirOwningReaders() {
            WordEvidenceCorpusManifest manifest = WordEvidenceCorpusManifestLoader.Load();
            foreach (WordEvidenceCorpusArtifact artifact in manifest.Artifacts.Where(item => item.Path != null)) {
                string path = WordEvidenceCorpusManifestLoader.ResolveArtifactPath(artifact);
                switch (Path.GetExtension(path).ToLowerInvariant()) {
                    case ".docx":
                        using (WordDocument document = WordDocument.Load(path)) {
                            Assert.NotNull(document._wordprocessingDocument.MainDocumentPart);
                            Assert.NotNull(document.InspectFeatures());
                        }
                        break;
                    case ".doc":
                        using (LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(path)) {
                            result.EnsureNoImportErrors();
                            Assert.True(result.HasDocument);
                        }
                        break;
                    case ".html":
                        string html = File.ReadAllText(path);
                        Assert.Contains("<ruby>", html, StringComparison.OrdinalIgnoreCase);
                        Assert.Contains("aria-description", html, StringComparison.OrdinalIgnoreCase);
                        break;
                    case ".json":
                        using (JsonDocument json = JsonDocument.Parse(File.ReadAllText(path))) {
                            Assert.Equal(1, json.RootElement.GetProperty("schemaVersion").GetInt32());
                        }
                        break;
                    default:
                        throw new InvalidDataException("Unexpected Word evidence artifact: " + path);
                }
            }
        }
    }
}
