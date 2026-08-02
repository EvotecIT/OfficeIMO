using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class ExternalEmailCorpusTests {
    [ProducerCorpusFact("MSGReader")]
    public void ProcessesAllMsgReaderSamplesWhenCorpusIsAvailable() {
        string? repository = ExternalEmailCorpusHarness.FindRepository("MSGReader");
        Assert.NotNull(repository);

        ExternalCorpusResult result = ExternalEmailCorpusHarness.RunMsgReader(repository);

        Assert.True(result.ApplicableArtifacts > 0, "No applicable MsgReader artifacts were found.");
        Assert.True(result.Failures.Count == 0, result.FormatFailures());
        Assert.Equal(result.CandidateArtifacts, result.ApplicableArtifacts);
        Assert.Equal(0, result.SkippedArtifacts);
        Assert.Equal(result.CandidateArtifacts, result.ArtifactSha256.Count);
        Assert.Equal(result.CandidateArtifacts, result.ArtifactEvidence.Count);
        Assert.All(result.ArtifactEvidence, evidence => {
            Assert.False(Path.IsPathRooted(evidence.RelativePath));
            Assert.Equal(64, evidence.Sha256.Length);
            Assert.True(evidence.MatchedExpectedSemantics);
        });
    }

    [ProducerCorpusFact("MimeKit")]
    public void ProcessesMimeKitMimeTnefAndMboxCorporaWhenAvailable() {
        string? repository = ExternalEmailCorpusHarness.FindRepository("MimeKit");
        Assert.NotNull(repository);

        ExternalCorpusResult result = ExternalEmailCorpusHarness.RunMimeKit(repository);

        Assert.True(result.ApplicableArtifacts > 0, "No applicable MimeKit artifacts were found.");
        Assert.True(result.Failures.Count == 0, result.FormatFailures());
        Assert.Equal(result.CandidateArtifacts, result.ApplicableArtifacts);
        Assert.Equal(0, result.SkippedArtifacts);
        Assert.Equal(result.CandidateArtifacts, result.ArtifactSha256.Count);
        Assert.Equal(result.CandidateArtifacts, result.ArtifactEvidence.Count);
        Assert.All(result.ArtifactEvidence, evidence => {
            Assert.False(Path.IsPathRooted(evidence.RelativePath));
            Assert.Equal(64, evidence.Sha256.Length);
            Assert.True(evidence.MatchedExpectedSemantics);
        });
    }
}

public sealed class ProducerCorpusFactAttribute : FactAttribute {
    public ProducerCorpusFactAttribute(string repositoryName) {
        if (ExternalEmailCorpusHarness.FindRepository(repositoryName) == null) {
            Skip = "Set OFFICEIMO_EMAIL_CORPUS_ROOT or EVOTEC_GITHUB_ROOT to a root containing " +
                repositoryName + " to run this producer-corpus test.";
        }
    }
}
