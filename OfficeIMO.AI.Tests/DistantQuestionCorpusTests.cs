using OfficeIMO.AI;
using Xunit;

namespace OfficeIMO.AI.Tests;

public sealed class DistantQuestionCorpusTests {
    [Fact]
    public async Task ComparisonInputPreservesItsDeclaredDistantPagesThroughTheRealPdfReader() {
        EvaluationCase item = EvaluationChallengeCorpus.Create().Single(item => item.Id == "distant-question-development");
        OfficeAiDocument document = await DocumentInputs.ReadAsync(item.Source, "comparison.pdf", false,
            item.Request.Pages, item.Request.Limits, CancellationToken.None);
        Assert.Equal(Enumerable.Range(1, 8), document.Pages);
        Assert.Contains(document.Evidence, item => item.Page == 1 && item.Text.Contains("Northern depot: 29 crates.", StringComparison.Ordinal));
        Assert.Contains(document.Evidence, item => item.Page == 8 && item.Text.Contains("Southern depot: 11 crates.", StringComparison.Ordinal));
        Assert.False(document.HasSourceDiagnostics);
    }
}
