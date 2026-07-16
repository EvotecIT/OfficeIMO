using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class ExternalEmailCorpusTests {
    [Fact]
    public void ProcessesAllMsgReaderSamplesWhenCorpusIsAvailable() {
        string? repository = ExternalEmailCorpusHarness.FindRepository("MSGReader");
        if (repository == null) return;

        ExternalCorpusResult result = ExternalEmailCorpusHarness.RunMsgReader(repository);

        Assert.True(result.ApplicableArtifacts > 0, "No applicable MsgReader artifacts were found.");
        Assert.True(result.Failures.Count == 0, result.FormatFailures());
        Assert.Equal(result.CandidateArtifacts, result.ApplicableArtifacts);
        Assert.Equal(0, result.SkippedArtifacts);
    }

    [Fact]
    public void ProcessesMimeKitMimeTnefAndMboxCorporaWhenAvailable() {
        string? repository = ExternalEmailCorpusHarness.FindRepository("MimeKit");
        if (repository == null) return;

        ExternalCorpusResult result = ExternalEmailCorpusHarness.RunMimeKit(repository);

        Assert.True(result.ApplicableArtifacts > 0, "No applicable MimeKit artifacts were found.");
        Assert.True(result.Failures.Count == 0, result.FormatFailures());
        Assert.Equal(result.CandidateArtifacts, result.ApplicableArtifacts);
        Assert.Equal(0, result.SkippedArtifacts);
    }
}
