using System.Text.Json;
using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public sealed class Markdown_CommonMark_Examples_Tests {
    private static readonly Lazy<IReadOnlyList<CommonMarkExampleFixture>> CachedExamples = new(LoadExamples);

    public static IEnumerable<object[]> CommonMarkSmokeExamples() {
        return CachedExamples.Value.Select(example => new object[] {
            $"example {example.Example} ({example.Section})",
            example
        });
    }

    [Theory]
    [MemberData(nameof(CommonMarkSmokeExamples))]
    public void CommonMark_Profile_Matches_Official_Examples_And_Keeps_SyntaxTree_Invariants(string _, CommonMarkExampleFixture example) {
        var result = MarkdownReader.ParseWithSyntaxTree(example.Markdown, MarkdownReaderOptions.CreateCommonMarkProfile());
        var html = result.Document.ToHtmlFragment(CommonMarkHtmlComparison.CreatePlainHtmlOptions());

        Assert.Equal(CommonMarkHtmlComparison.Normalize(example.Html), CommonMarkHtmlComparison.Normalize(html));
        Assert.Equal(example.TopLevelKinds, result.SyntaxTree.Children.Select(node => node.Kind.ToString()).ToArray());
        MarkdownSpecSyntaxAssert.AssertSyntaxAssertions(result.SyntaxTree, example.SyntaxAssertions);
        MarkdownSpecSyntaxAssert.AssertSyntaxAssertions(result.FinalSyntaxTree, example.FinalSyntaxAssertions);

        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.SemanticTreeIsWellFormed(result.Document);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    private static IReadOnlyList<CommonMarkExampleFixture> LoadExamples() {
        string path = Path.Combine(
            AppContext.BaseDirectory,
            "..", "..", "..",
            "Markdown",
            "Fixtures",
            "CommonMark",
            "commonmark-0.31.2-smoke.json");

        string json = File.ReadAllText(path);
        var examples = JsonSerializer.Deserialize<List<CommonMarkExampleFixture>>(json, new JsonSerializerOptions {
            PropertyNameCaseInsensitive = true
        });

        Assert.NotNull(examples);
        Assert.NotEmpty(examples);
        Assert.All(examples!, example => Assert.NotEmpty(example.TopLevelKinds));
        return examples!;
    }

    public sealed class CommonMarkExampleFixture {
        public int Example { get; set; }
        public string Section { get; set; } = string.Empty;
        public string Markdown { get; set; } = string.Empty;
        public string Html { get; set; } = string.Empty;
        public string[] TopLevelKinds { get; set; } = [];
        public MarkdownSpecSyntaxAssertionFixture[] SyntaxAssertions { get; set; } = [];
        public MarkdownSpecSyntaxAssertionFixture[] FinalSyntaxAssertions { get; set; } = [];
    }
}
