using System.Text.Json;
using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public sealed class Markdown_GitHubFlavoredMarkdown_Examples_Tests {
    private static readonly Lazy<IReadOnlyList<GfmExampleFixture>> CachedExamples = new(LoadExamples);

    public static IEnumerable<object[]> GfmSmokeExamples() {
        return CachedExamples.Value.Select(example => new object[] {
            $"{example.Section} ({example.Source})",
            example
        });
    }

    [Theory]
    [MemberData(nameof(GfmSmokeExamples))]
    public void Gfm_Profile_Matches_Official_Smoke_Examples_And_Keeps_SyntaxTree_Invariants(string _, GfmExampleFixture example) {
        var result = MarkdownReader.ParseWithSyntaxTree(example.Markdown, MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile());
        var html = result.Document.ToHtmlFragment(GfmHtmlComparison.CreatePlainHtmlOptions());

        Assert.Equal(GfmHtmlComparison.Normalize(example.Html), GfmHtmlComparison.Normalize(html));
        Assert.Equal(example.TopLevelKinds, result.SyntaxTree.Children.Select(node => node.Kind.ToString()).ToArray());
        MarkdownSpecSyntaxAssert.AssertSyntaxAssertions(result.SyntaxTree, example.SyntaxAssertions);

        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.SemanticTreeIsWellFormed(result.Document);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    private static IReadOnlyList<GfmExampleFixture> LoadExamples() {
        string path = Path.Combine(
            AppContext.BaseDirectory,
            "..", "..", "..",
            "Markdown",
            "Fixtures",
            "GitHubFlavoredMarkdown",
            "cmark-gfm-extensions-smoke.json");

        string json = File.ReadAllText(path);
        var examples = JsonSerializer.Deserialize<List<GfmExampleFixture>>(json, new JsonSerializerOptions {
            PropertyNameCaseInsensitive = true
        });

        Assert.NotNull(examples);
        Assert.NotEmpty(examples);
        Assert.All(examples!, example => Assert.NotEmpty(example.TopLevelKinds));
        return examples!;
    }
}
