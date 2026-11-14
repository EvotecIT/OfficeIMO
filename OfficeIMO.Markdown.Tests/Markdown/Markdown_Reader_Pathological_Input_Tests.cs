using System.Runtime.ExceptionServices;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Markdown;
using Xunit;
using static OfficeIMO.Markdown.MarkdownLoadOptions;

namespace OfficeIMO.Tests.MarkdownSuite;

// A stack overflow ends the test host, so these cases run on a typical 1 MB worker stack. The
// watchdog only catches hangs; it is far above the time the cases need.
public sealed class Markdown_Reader_Pathological_Input_Tests {
    private const int WorkerStackBytes = 1024 * 1024;
    private static readonly TimeSpan Watchdog = TimeSpan.FromSeconds(60);

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void DeepInlineNestingIsRejectedWithTheNestingLimit(bool syntaxTree) {
        // 300 pairs used to overflow the stack and end the process.
        string markdown = NestedEmphasis(300);
        var options = MarkdownLoadOptions.CreateCommonMarkProfile();
        options.CaptureSyntaxTree = syntaxTree;

        var exception = RunOnWorkerStack(() => Assert.Throws<InvalidDataException>(
            () => MarkdownDocument.Parse(markdown, options)));

        Assert.Contains("inline nesting exceeds MaxNestingDepth (64)", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void InlineNestingAtTheLimitParsesAndRenders() {
        // 32 pairs are 64 nested inline containers, the deepest a default parse accepts.
        string markdown = NestedEmphasis(32);

        var (html, written) = RunOnWorkerStack(() => {
            var document = MarkdownReader.ParseWithSyntaxTree(markdown).Document;
            return (document.ToHtml(CommonMarkHtmlComparison.CreatePlainHtmlOptions()), document.ToMarkdown());
        });

        Assert.Matches("(<em>a <strong>a ){32}b( a</strong> a</em>){32}", html);
        Assert.Equal(markdown, written.Trim());
    }

    [Fact]
    public void HandBuiltTreesDeeperThanTheStackFailWithACatchableException() {
        InlineSequence inlines = new InlineSequence().Text("x");
        for (int level = 0; level < 100_000; level++) {
            var parent = new InlineSequence();
            parent.ReplaceItems(new IMarkdownInline[] { new BoldSequenceInline(inlines) });
            inlines = parent;
        }

        var document = MarkdownDocument.Create();
        document.Add(new ParagraphBlock(inlines));

        RunOnWorkerStack(() => Assert.Throws<InsufficientExecutionStackException>(() => document.ToHtml()));
        RunOnWorkerStack(() => Assert.Throws<InsufficientExecutionStackException>(() => document.ToMarkdown()));
    }

    // Ported from cmark-gfm test/pathological_tests.py at the commit recorded in
    // Fixtures/GitHubFlavoredMarkdown/cmark-gfm-provenance.json (sha256 b200aa0f...af69).
    // Expected patterns are upstream's with line breaks dropped, because both sides are compared
    // without them. A null pattern means the case only has to complete: the output differs from
    // cmark-gfm for a reason unrelated to resource use, noted beside the case.
    private static readonly Dictionary<string, Func<PathologicalCase>> Cases = BuildCases();

    public static IEnumerable<object[]> CmarkGfmPathologicalCases() =>
        Cases.Keys.Select(static name => new object[] { name });

    private static Dictionary<string, Func<PathologicalCase>> BuildCases() {
        var cases = new Dictionary<string, Func<PathologicalCase>>(StringComparer.Ordinal);
        void Add(string name, Func<PathologicalCase> create) => cases.Add(name, create);
        Add("nested strong emph", () => Case(Repeat("*a **a ", 65000) + "b" + Repeat(" a** a*", 65000), rejectedByNestingLimit: true));
        Add("many emph closers with no openers", () => Case(Repeat("a_ ", 65000), "(a[_] ){64999}a_"));
        Add("many emph openers with no closers", () => Case(Repeat("_a ", 65000), "(_a ){64999}_a"));
        Add("many link closers with no openers", () => Case(Repeat("a]", 65000), @"(a\]){65000}"));
        Add("many link openers with no closers", () => Case(Repeat("[a", 65000), @"(\[a){65000}"));
        Add("mismatched openers and closers", () => Case(Repeat("*a_ ", 50000), "([*]a[_] ){49999}[*]a_"));
        Add("openers and closers multiple of 3", () => Case("a**b" + Repeat("c* ", 50000), "a[*][*]b(c[*] ){49999}c[*]"));
        Add("link openers and emph closers", () => Case(Repeat("[ a_", 50000), @"(\[ a_){50000}"));
        Add("pattern [ (]( repeated", () => Case(Repeat("[ (](", 80000), @"(\[ \(\]\(){80000}"));
        Add("pattern ![[]() repeated", () => Case(Repeat("![[]()", 160000), "(!\\[<a href=\"\"></a>){160000}"));
        Add("hard link/emph case", () => Case("**x [a*b**c*](d)", "\\*\\*x <a href=\"d\">a<em>b\\*\\*c</em></a>"));
        Add("nested brackets", () => Case(new string('[', 50000) + "a" + new string(']', 50000), @"\[{50000}a\]{50000}"));
        // Block containers are limited by MaxNestingDepth as well.
        Add("nested block quotes", () => Case(Repeat("> ", 50000) + "a", rejectedByNestingLimit: true));
        Add("deeply nested lists", () => Case(string.Concat(Enumerable.Range(0, 1000).Select(level => new string(' ', 2 * level) + "* a\n")), rejectedByNestingLimit: true));
        // U+0000 becomes a space rather than U+FFFD.
        Add("U+0000 in input", () => Case("abc\u0000de\u0000", pattern: null));
        Add("backticks", () => Case(string.Concat(Enumerable.Range(1, 4999).Select(length => "e" + new string('`', length))), "^<p>[e`]*</p>$"));
        Add("unclosed links A", () => Case(Repeat("[a](<b", 30000), @"(\[a\]\(&lt;b){30000}"));
        Add("unclosed links B", () => Case(Repeat("[a](b", 30000), @"(\[a\]\(b){30000}"));
        Add("unclosed <!--", () => Case("</" + Repeat("<!--", 300000), @"&lt;/(&lt;!--){300000}"));
        // A bare carriage return is not treated as a line ending, so no table is formed.
        Add("tables", () => Case(Repeat("aaa\rbbb\n-\v\n", 30000), pattern: null, profile: MarkdownDialectProfile.GitHubFlavoredMarkdown));
        Add("reference collisions", ReferenceCollisionsCase);
        // Inputs that were quadratic in the reader before the scans were indexed.
        Add("unclosed link openers with spaces", () => Case(Repeat("[a ", 65000), @"(\[a ){64999}"));
        Add("unclosed html tag starts", () => Case(Repeat("<a ", 65000), "(&lt;a ){64999}"));
        Add("unmatched backtick runs", () => Case(Repeat("`a``", 32500), "^<p>"));
        // Every opener's label holds all the inner links, so labels were copied and rescanned per opener.
        Add("nested inline links", () => Case(Repeat("[", 65000) + "a" + Repeat("](u)", 65000), "^<p>\\[{64999}<a href=\"u\">a</a>(\\]\\(u\\)){64999}</p>$"));
        Add("nested inline links in emphasis", () => Case(Repeat("*[", 65000) + "a" + Repeat("](u)*", 65000), "^<p>(<em>|</em>|\\[)+<a href=\"u\">a</a>(<em>|</em>|\\]\\(u\\))+</p>$"));
        Add("nested inline links around raw html", () => Case(Repeat("[<b>", 65000) + "a" + Repeat("</b>](u)", 65000), "^<p>(\\[<b>){64999}<a href=\"u\"><b>a</b></a>(</b>\\]\\(u\\)){64999}</p>$"));
        Add("nested reference links", () => Case("[r]: /x\n\n" + Repeat("[", 65000) + "a" + Repeat("][r]", 65000), "^<p>\\[{64999}<a href=\"/x\">a</a>(\\]<a href=\"/x\">r</a>){64999}</p>$"));
        // Each nested container started its parse from a copy of every definition.
        Add("many definitions and block quotes", () => Case(
            string.Concat(Enumerable.Range(0, 40000).Select(index => "[ref" + index + "]: /u" + index + "\n")) + "\n" +
            string.Concat(Enumerable.Range(0, 40000).Select(index => "> quote [ref" + index + "]\n> - item\n\n")),
            "<blockquote><p>quote <a href=\"/u39999\">ref39999</a></p><ul><li>item</li></ul></blockquote>$"));
        // The default profile checks each bare autolink candidate against its line.
        Add("bare autolink candidates on one line", () => Case(Repeat("www.a ", 65000), "^<p>", profile: MarkdownDialectProfile.OfficeIMO));
        return cases;
    }

    [Theory]
    [MemberData(nameof(CmarkGfmPathologicalCases))]
    public void PathologicalInputCompletes(string name) {
        PathologicalCase pathological = Cases[name]();
        var options = MarkdownLoadOptions.CreateProfile(pathological.Profile);

        if (pathological.RejectedByNestingLimit) {
            RunOnWorkerStack(() => Assert.Throws<InvalidDataException>(() => MarkdownDocument.Parse(pathological.Markdown, options)));
            return;
        }

        string html = RunOnWorkerStack(() => MarkdownDocument.Parse(pathological.Markdown, options)
            .ToHtml(CommonMarkHtmlComparison.CreatePlainHtmlOptions()));

        Assert.NotEmpty(html);
        if (pathological.Pattern != null) {
            Assert.Matches(new Regex(pathological.Pattern), html.Replace("\n", string.Empty).Trim());
        }
    }

    public sealed class PathologicalCase {
        public PathologicalCase(string markdown, string? pattern, bool rejectedByNestingLimit, MarkdownDialectProfile profile) {
            Markdown = markdown;
            Pattern = pattern;
            RejectedByNestingLimit = rejectedByNestingLimit;
            Profile = profile;
        }

        public string Markdown { get; }
        public string? Pattern { get; }
        public bool RejectedByNestingLimit { get; }
        public MarkdownDialectProfile Profile { get; }
    }

    private static PathologicalCase Case(string markdown, string? pattern = null, bool rejectedByNestingLimit = false, MarkdownDialectProfile profile = MarkdownDialectProfile.CommonMark) =>
        new PathologicalCase(markdown, pattern, rejectedByNestingLimit, profile);

    private static PathologicalCase ReferenceCollisionsCase() {
        // Keys whose cmark reference-map hash collides, so a hashed lookup degrades to a list scan.
        const int count = 50000;
        var keys = new List<string>(count);
        for (int index = 0; keys.Count < count; index++) {
            string key = "x" + index.ToString(System.Globalization.CultureInfo.InvariantCulture);
            uint hash = 0;
            foreach (char character in key) {
                hash = character + (hash << 6) + (hash << 16) - hash;
            }

            if (hash % 16 == 0) {
                keys.Add(key);
            }
        }

        string badKey = keys[0];
        var markdown = new StringBuilder();
        foreach (string key in keys.Skip(1)) {
            markdown.Append('[').Append(key).Append("]: /url\n\n[").Append(badKey).Append("]\n\n");
        }

        return Case(markdown.ToString(), "(<p>\\[" + badKey + "\\]</p>){" + (count - 1) + "}");
    }

    private static string NestedEmphasis(int pairs) =>
        Repeat("*a **a ", pairs) + "b" + Repeat(" a** a*", pairs);

    private static string Repeat(string value, int count) =>
        string.Concat(Enumerable.Repeat(value, count));

    private static T RunOnWorkerStack<T>(Func<T> action) {
        T result = default!;
        ExceptionDispatchInfo? failure = null;
        var thread = new Thread(() => {
            try {
                result = action();
            } catch (Exception exception) {
                failure = ExceptionDispatchInfo.Capture(exception);
            }
        }, WorkerStackBytes) {
            IsBackground = true
        };
        thread.Start();
        Assert.True(thread.Join(Watchdog), "Parsing did not complete within " + Watchdog + ".");
        failure?.Throw();
        return result;
    }
}
