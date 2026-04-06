using System.Text.Json;
using System.Text.RegularExpressions;
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
        var html = result.Document.ToHtmlFragment(CreatePlainHtmlOptions());

        Assert.Equal(NormalizeHtmlForComparison(example.Html), NormalizeHtmlForComparison(html));
        Assert.Equal(example.TopLevelKinds, result.SyntaxTree.Children.Select(node => node.Kind.ToString()).ToArray());
        MarkdownSpecSyntaxAssert.AssertSyntaxAssertions(result.SyntaxTree, example.SyntaxAssertions);
        MarkdownSpecSyntaxAssert.AssertSyntaxAssertions(result.FinalSyntaxTree, example.FinalSyntaxAssertions);

        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.SemanticTreeIsWellFormed(result.Document);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    private static HtmlOptions CreatePlainHtmlOptions() {
        return new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null
        };
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

    private static string NormalizeHtmlForComparison(string html) {
        if (string.IsNullOrWhiteSpace(html)) {
            return string.Empty;
        }

        var sb = new StringBuilder(html.Length);
        bool inTag = false;
        bool lastWasWhitespace = false;

        for (int i = 0; i < html.Length; i++) {
            char ch = html[i];
            if (ch == '<') {
                if (!inTag && lastWasWhitespace && sb.Length > 0 && sb[sb.Length - 1] != '>') {
                    sb.Append(' ');
                }

                inTag = true;
                lastWasWhitespace = false;
                sb.Append(ch);
                continue;
            }

            if (ch == '>') {
                inTag = false;
                lastWasWhitespace = false;
                sb.Append(ch);
                continue;
            }

            if (inTag) {
                sb.Append(ch);
                continue;
            }

            if (char.IsWhiteSpace(ch)) {
                lastWasWhitespace = true;
                continue;
            }

            if (lastWasWhitespace && sb.Length > 0 && sb[sb.Length - 1] != '>') {
                sb.Append(' ');
            }

            lastWasWhitespace = false;
            sb.Append(ch);
        }

        string normalized = sb.ToString()
            .Replace("> <", "><")
            .Replace("&#39;", "'")
            .Replace("&#x27;", "'");
        normalized = Regex.Replace(normalized, "<h([1-6])\\s+id=\"[^\"]*\">", "<h$1>", RegexOptions.CultureInvariant);
        normalized = normalized
            .Replace(" <ul", "<ul")
            .Replace(" <ol", "<ol")
            .Replace(" <blockquote", "<blockquote")
            .Replace(" <pre", "<pre")
            .Replace(" <table", "<table")
            .Replace(" <p", "<p");

        return normalized.Trim();
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
