using OfficeIMO.Yaml.Tests.Conformance;
using Xunit;

namespace OfficeIMO.Yaml.Tests;

// The pinned suite and the corpus are the reference every YAML engine is measured against.
// Loading fails when files drift from the case count or SHA-256 values their manifests record.
public class YamlTestAssetTests {
    [Fact]
    public void SuiteCasesMatchTheManifest() {
        Assert.NotEmpty(YamlTestSuite.Load().Cases);
    }

    [Fact]
    public void CorpusFilesMatchTheirRecordedHashes() {
        Assert.NotEmpty(YamlCorpus.Load());
    }
}
