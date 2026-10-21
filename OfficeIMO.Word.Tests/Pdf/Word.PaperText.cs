using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordPaperTextTests {
    [Theory]
    [InlineData("ShortTerm", "Short term")]
    [InlineData("AnalystOverlay", "Analyst overlay")]
    [InlineData("DNSHealth", "DNS health")]
    [InlineData("NotApplicable", "Not applicable")]
    public void HumanizeIdentifier_EnumNames_UsesSentenceCaseAndPreservesAcronyms(string source, string expected)
        => Assert.Equal(expected, WordPaperText.HumanizeIdentifier(source));

    [Fact]
    public void SplitIdentifier_PropertyLabels_OffersWordsWithoutSplittingAcronyms()
        => Assert.Equal("DNS Enrollment Rights", WordPaperText.SplitIdentifier("DNSEnrollmentRights"));

    [Theory]
    [InlineData("DNS: Zone Transfer Posture For ExampleApp", "DNS: zone transfer posture for ExampleApp")]
    [InlineData("iSCSI Target Configuration", "iSCSI target configuration")]
    public void SentenceCase_OwnedTitle_PreservesTechnicalTerms(string source, string expected)
        => Assert.Equal(expected, WordPaperText.SentenceCase(source));
}
