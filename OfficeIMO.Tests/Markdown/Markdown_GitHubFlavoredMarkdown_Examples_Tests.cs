using System.Text.Json;
using System.Text.RegularExpressions;
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
        var html = result.Document.ToHtmlFragment(CreatePlainHtmlOptions());

        Assert.Equal(NormalizeHtmlForComparison(example.Html), NormalizeHtmlForComparison(html));
        Assert.Equal(example.TopLevelKinds, result.SyntaxTree.Children.Select(node => node.Kind.ToString()).ToArray());
        MarkdownSpecSyntaxAssert.AssertSyntaxAssertions(result.SyntaxTree, example.SyntaxAssertions);

        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.SemanticTreeIsWellFormed(result.Document);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    private static HtmlOptions CreatePlainHtmlOptions() {
        return new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            GitHubTaskListHtml = true,
            GitHubFootnoteHtml = true
        };
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

    private static string NormalizeHtmlForComparison(string html) {
        if (string.IsNullOrWhiteSpace(html)) {
            return string.Empty;
        }

        html = Regex.Replace(
            html,
            "<input([^>]*)>",
            static match => {
                string attrs = match.Groups[1].Value;
                bool isChecked = Regex.IsMatch(attrs, "\\bchecked(?:\\s*=\\s*\"[^\"]*\")?", RegexOptions.CultureInvariant);
                bool isDisabled = Regex.IsMatch(attrs, "\\bdisabled(?:\\s*=\\s*\"[^\"]*\")?", RegexOptions.CultureInvariant);
                return $"<input type=\"checkbox\"{(isChecked ? " checked=\"\"" : string.Empty)}{(isDisabled ? " disabled=\"\"" : string.Empty)} />";
            },
            RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

        html = html
            .Replace(" class=\"contains-task-list\"", string.Empty)
            .Replace(" class=\"task-list-item\"", string.Empty)
            .Replace(" class=\"task-list-item-checkbox\"", string.Empty);

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
        return normalized.Trim();
    }

    public sealed class GfmExampleFixture {
        public string Source { get; set; } = string.Empty;
        public string Section { get; set; } = string.Empty;
        public string Markdown { get; set; } = string.Empty;
        public string Html { get; set; } = string.Empty;
        public string[] TopLevelKinds { get; set; } = [];
        public MarkdownSpecSyntaxAssertionFixture[] SyntaxAssertions { get; set; } = [];
    }
}
