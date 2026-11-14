using OfficeIMO.Tool.Agent;
using Xunit;

namespace OfficeIMO.Tool.Tests;

public sealed class AgentRankedSearchTests {
    private const string Markdown =
        "# Travel\n\nBook travel through the portal. Revenue rules are separate.\n\n" +
        "# Revenue\n\nRevenue grew. Subscription revenue and services revenue rose.\n\n" +
        "# Notes\n\nArrows like -> and => appear in the appendix.\n";

    [Fact]
    public async Task DocumentSearchRanksHitsAndPagesInRankOrder() {
        string path = Path.Combine(Path.GetTempPath(), "officeimo-agent-ranked-" + Guid.NewGuid().ToString("N") + ".md");
        try {
            await File.WriteAllTextAsync(path, Markdown);
            var service = new OfficeImoAgentService();

            AgentSearchResult all = await service.SearchAsync(path, "REVENUE", take: 10);
            Assert.True(all.Results.Count >= 2);
            Assert.All(all.Results, hit => Assert.NotNull(hit.Score));
            Assert.Equal(all.Results.Select(static hit => hit.Score!.Value).OrderByDescending(static score => score),
                all.Results.Select(static hit => hit.Score!.Value));
            Assert.Contains("Subscription revenue", all.Results[0].Snippet, StringComparison.Ordinal);

            AgentSearchResult first = await service.SearchAsync(path, "revenue", take: 1);
            Assert.Equal(1, first.NextCursor);
            AgentSearchResult second = await service.SearchAsync(path, "revenue", take: 1, cursor: 1);
            Assert.Equal(all.Results[0].Id, Assert.Single(first.Results).Id);
            Assert.Equal(all.Results[1].Id, Assert.Single(second.Results).Id);

            AgentSearchHit phrase = Assert.Single((await service.SearchAsync(path, "\"services revenue\"")).Results);
            Assert.Equal(all.Results[0].Id, phrase.Id);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public async Task QueriesWithoutWordsKeepLiteralMatching() {
        string path = Path.Combine(Path.GetTempPath(), "officeimo-agent-literal-" + Guid.NewGuid().ToString("N") + ".md");
        try {
            await File.WriteAllTextAsync(path, Markdown);
            var service = new OfficeImoAgentService();

            AgentSearchHit arrow = Assert.Single((await service.SearchAsync(path, "->")).Results);
            Assert.Null(arrow.Score);
            Assert.Contains("->", arrow.Snippet, StringComparison.Ordinal);

            AgentSearchResult fragment = await service.SearchAsync(path, "ubscript");
            Assert.All(fragment.Results, hit => Assert.Null(hit.Score));
            Assert.NotEmpty(fragment.Results);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
