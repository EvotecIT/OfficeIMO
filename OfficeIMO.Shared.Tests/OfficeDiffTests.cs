using System.Text;
using OfficeIMO.Diff;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class OfficeDiffTests {
    public static IEnumerable<object[]> Algorithms() {
        yield return new object[] { OfficeDiffAlgorithm.Myers };
        yield return new object[] { OfficeDiffAlgorithm.Patience };
        yield return new object[] { OfficeDiffAlgorithm.Histogram };
    }

    [Theory]
    [MemberData(nameof(Algorithms))]
    public void EditScriptTransformsOldIntoNewAndBack(OfficeDiffAlgorithm algorithm) {
        var random = new Random(1234 + (int)algorithm);
        for (int round = 0; round < 600; round++) {
            int[] oldItems = RandomSequence(random, random.Next(0, 40), 1 + random.Next(6));
            int[] newItems = Mutate(random, oldItems, 1 + random.Next(6));
            var options = new OfficeDiffOptions { Algorithm = algorithm, SemanticCleanup = round % 3 == 0 };

            OfficeDiffResult<int> diff = OfficeDiff.Compare(oldItems, newItems, options: options);

            Assert.Equal(newItems, diff.Apply(oldItems));
            Assert.Equal(oldItems, diff.Revert(newItems));
            AssertCoversBothSides(diff);
        }
    }

    [Fact]
    public void MyersEditDistanceIsMinimal() {
        var random = new Random(42);
        for (int round = 0; round < 800; round++) {
            int[] oldItems = RandomSequence(random, random.Next(0, 60), 1 + random.Next(5));
            int[] newItems = round % 2 == 0
                ? Mutate(random, oldItems, 1 + random.Next(5))
                : RandomSequence(random, random.Next(0, 60), 1 + random.Next(5));

            OfficeDiffResult<int> diff = OfficeDiff.Compare(oldItems, newItems);

            int lcs = ReferenceLcsLength(oldItems, newItems);
            Assert.False(diff.LimitReached);
            Assert.Equal(oldItems.Length + newItems.Length - (2 * lcs), diff.EditDistance);
        }
    }

    [Theory]
    [MemberData(nameof(Algorithms))]
    public void LargeInputsWithFewChangesStayFastAndExact(OfficeDiffAlgorithm algorithm) {
        // 40,000 x 40,000 items is 1.6 billion DP cells; the linear-space search only explores the few edits.
        string[] oldLines = Enumerable.Range(0, 40_000).Select(static index => "line " + index).ToArray();
        var newLines = oldLines.ToList();
        newLines[100] = "changed";
        newLines.Insert(20_000, "inserted");
        newLines.RemoveAt(35_000);

        OfficeDiffResult<string> diff = OfficeDiff.Compare(oldLines, newLines, StringComparer.Ordinal, new OfficeDiffOptions { Algorithm = algorithm });

        Assert.False(diff.LimitReached);
        Assert.Equal(2, diff.DeletedCount);
        Assert.Equal(2, diff.InsertedCount);
        Assert.Equal(newLines, diff.Apply(oldLines));
    }

    [Fact]
    public void WorkLimitStillReturnsACompleteScript() {
        var random = new Random(7);
        int[] oldItems = RandomSequence(random, 3000, 4);
        int[] newItems = RandomSequence(random, 3000, 4);

        OfficeDiffResult<int> diff = OfficeDiff.Compare(oldItems, newItems, options: new OfficeDiffOptions { MaxWorkUnits = 1000 });

        Assert.True(diff.LimitReached);
        Assert.Equal(newItems, diff.Apply(oldItems));
        Assert.Equal(oldItems, diff.Revert(newItems));
    }

    [Fact]
    public void ZeroTimeoutStopsAtOnceWithACompleteScript() {
        var random = new Random(8);
        int[] oldItems = RandomSequence(random, 2000, 3);
        int[] newItems = RandomSequence(random, 2000, 3);

        OfficeDiffResult<int> diff = OfficeDiff.Compare(oldItems, newItems, options: new OfficeDiffOptions { Timeout = TimeSpan.Zero });

        Assert.True(diff.LimitReached);
        Assert.Equal(newItems, diff.Apply(oldItems));
    }

    [Fact]
    public void ApplyRejectsADifferentSource() {
        OfficeDiffResult<char> diff = OfficeDiff.Compare("abc".ToCharArray(), "abd".ToCharArray());

        Assert.Throws<InvalidOperationException>(() => diff.Apply("xbc".ToCharArray()));
        Assert.Throws<InvalidOperationException>(() => diff.Revert("abdd".ToCharArray()));
    }

    [Fact]
    public void ComparerDecidesEquality() {
        string[] oldItems = { "Alpha", "beta", "GAMMA" };
        string[] newItems = { "alpha", "BETA", "gamma" };

        Assert.True(OfficeDiff.Compare(oldItems, newItems, StringComparer.OrdinalIgnoreCase).AreEqual);
        Assert.Equal(6, OfficeDiff.Compare(oldItems, newItems, StringComparer.Ordinal).EditDistance);
    }

    [Fact]
    public void NullItemsAreCompared() {
        string?[] oldItems = { "a", null, "b" };
        string?[] newItems = { null, "b" };

        OfficeDiffResult<string?> diff = OfficeDiff.Compare(oldItems, newItems);

        Assert.Equal(1, diff.EditDistance);
        Assert.Equal(newItems, diff.Apply(oldItems));
    }

    [Fact]
    public void PatienceAlignsOnUniqueLinesAroundMovedBraces() {
        string oldText = "void A() {\n  a();\n}\n\nvoid B() {\n  b();\n}\n";
        string newText = "void B() {\n  b();\n}\n\nvoid A() {\n  a();\n}\n";

        OfficeDiffResult<OfficeTextToken> patience = OfficeDiff.CompareLines(oldText, newText, new OfficeDiffOptions { Algorithm = OfficeDiffAlgorithm.Patience });

        Assert.Equal(newText, Concat(patience.Apply(patience.Old)));
        Assert.Contains(patience.Edits, edit => edit.Operation == OfficeDiffOperation.Equal && patience.Old[edit.OldStart].Text.StartsWith("void", StringComparison.Ordinal));
    }

    [Fact]
    public void SemanticCleanupMergesShortEqualitiesBetweenChanges() {
        var options = new OfficeDiffOptions { SemanticCleanup = true };

        OfficeDiffResult<OfficeTextToken> diff = OfficeDiff.CompareWords("the quick brown fox", "a slow red fox", options);

        IReadOnlyList<OfficeDiffHunk> regions = diff.GetChangedRegions();
        OfficeDiffHunk region = Assert.Single(regions);
        Assert.Equal("the quick brown", Concat(diff.Old.Skip(region.OldStart).Take(region.OldLength)));
        Assert.Equal("a slow red", Concat(diff.New.Skip(region.NewStart).Take(region.NewLength)));
    }

    [Fact]
    public void WordChangesSlideToWholeSentences() {
        OfficeDiffResult<OfficeTextToken> diff = OfficeDiff.CompareWords("Keep this. Drop that. End.", "Keep this. End.");

        OfficeDiffEdit deletion = Assert.Single(diff.Edits, static edit => edit.Operation == OfficeDiffOperation.Delete);
        Assert.Equal("Drop that. ", Concat(diff.Old.Skip(deletion.OldStart).Take(deletion.OldLength)));
    }

    [Theory]
    [MemberData(nameof(Algorithms))]
    public void WordDiffsStayCorrectAndMinimalAfterBoundaryAlignment(OfficeDiffAlgorithm algorithm) {
        string[] vocabulary = { "a", "b", "c", ".", ",", " ", " ", "\n" };
        var random = new Random(99 + (int)algorithm);
        for (int round = 0; round < 400; round++) {
            string oldText = RandomText(random, vocabulary);
            string newText = random.Next(3) == 0 ? RandomText(random, vocabulary) : MutateText(random, oldText, vocabulary);

            OfficeDiffResult<OfficeTextToken> diff = OfficeDiff.CompareWords(oldText, newText, new OfficeDiffOptions { Algorithm = algorithm });

            Assert.Equal(newText, Concat(diff.Apply(diff.Old)));
            Assert.Equal(oldText, Concat(diff.Revert(diff.New)));
            AssertCoversBothSides(diff);
            if (algorithm == OfficeDiffAlgorithm.Myers) {
                int lcs = ReferenceLcsLength(diff.Old.Select(static token => token.Text).ToArray(), diff.New.Select(static token => token.Text).ToArray());
                Assert.Equal(diff.Old.Count + diff.New.Count - (2 * lcs), diff.EditDistance);
            }
        }
    }

    [Fact]
    public void TextOptionsIgnoreCaseAndWhitespace() {
        var options = new OfficeDiffOptions { IgnoreCase = true, IgnoreWhitespace = true };

        Assert.True(OfficeDiff.CompareWords("Hello   World", "hello world", options).AreEqual);
        Assert.True(OfficeDiff.CompareLines("a  b\nC\n", "a b\nc\n", options).AreEqual);
        Assert.False(OfficeDiff.CompareWords("Hello World", "hello world").AreEqual);
    }

    [Fact]
    public void TokenizersCoverTheSourceExactly() {
        string text = "Re\u0301sume\u0301 costs $5.00, \U0001F44D\U0001F3FD ok!\r\n東京は晴れ\nend";

        foreach (IReadOnlyList<OfficeTextToken> tokens in new[] { OfficeTextTokenizer.Lines(text), OfficeTextTokenizer.Words(text), OfficeTextTokenizer.Characters(text) }) {
            Assert.Equal(text, Concat(tokens));
            int offset = 0;
            foreach (OfficeTextToken token in tokens) {
                Assert.Equal(offset, token.Start);
                offset += token.Length;
            }
        }
    }

    [Fact]
    public void WordTokenizerKeepsGraphemesAndSplitsIdeographs() {
        IReadOnlyList<OfficeTextToken> tokens = OfficeTextTokenizer.Words("Re\u0301sume\u0301, \U0001F44D\U0001F3FD 東京");

        Assert.Equal(new[] { "Re\u0301sume\u0301", ",", " ", "\U0001F44D\U0001F3FD", " ", "東", "京" }, tokens.Select(static token => token.Text));
        Assert.Equal(OfficeTextTokenKind.Word, tokens[0].Kind);
        Assert.Equal(OfficeTextTokenKind.Punctuation, tokens[1].Kind);
        Assert.Equal(OfficeTextTokenKind.Whitespace, tokens[2].Kind);
    }

    [Fact]
    public void CharacterDiffNeverSplitsAGrapheme() {
        OfficeDiffResult<OfficeTextToken> diff = OfficeDiff.CompareCharacters("cafe\u0301", "cafe");

        OfficeDiffEdit deletion = Assert.Single(diff.Edits, static edit => edit.Operation == OfficeDiffOperation.Delete);
        Assert.Equal("e\u0301", diff.Old[deletion.OldStart].Text);
    }

    [Fact]
    public void UnifiedDiffMatchesTheGnuFormat() {
        string oldText = "one\ntwo\nthree\nfour\nfive\nsix\nseven\neight\nnine\nten\n";
        string newText = "one\ntwo\nTHREE\nfour\nfive\nsix\nseven\neight\nnine\nten\neleven";

        string patch = OfficeUnifiedDiff.Create(oldText, newText, "old.txt", "new.txt", contextLines: 2);

        Assert.Equal(
            "--- old.txt\n+++ new.txt\n" +
            "@@ -1,5 +1,5 @@\n one\n two\n-three\n+THREE\n four\n five\n" +
            "@@ -9,2 +9,3 @@\n nine\n ten\n+eleven\n\\ No newline at end of file\n",
            patch);
        Assert.Equal(string.Empty, OfficeUnifiedDiff.Create(oldText, oldText));
    }

    [Fact]
    public void UnifiedDiffWritesEmptyRangesAtTheLineBefore() {
        Assert.Equal("--- a\n+++ b\n@@ -0,0 +1 @@\n+x\n", OfficeUnifiedDiff.Create(string.Empty, "x\n"));
        Assert.Equal("--- a\n+++ b\n@@ -1 +0,0 @@\n-x\n", OfficeUnifiedDiff.Create("x\n", string.Empty));
    }

    [Fact]
    public void ChangeSetReportsKindsAnchorsAndMoves() {
        string oldText = "Alpha beta gamma delta epsilon zeta. Moved bit here. Iota kappa lambda mu nu. Old ending.";
        string newText = "Moved bit here. Alpha beta gamma delta epsilon zeta. Iota kappa lambda added words mu nu. New ending.";
        OfficeDiffResult<OfficeTextToken> diff = OfficeDiff.CompareWords(oldText, newText, new OfficeDiffOptions { SemanticCleanup = true });

        OfficeChangeSet changes = OfficeChangeSet.FromTextDiff(diff, new OfficeChangeSetOptions { OldSource = "old.txt", NewSource = "new.txt" });

        OfficeChange moved = Assert.Single(changes.Changes, static change => change.Kind == OfficeChangeKind.Moved);
        Assert.Equal("Moved bit here.", moved.OldText);
        Assert.Equal(moved.OldText, moved.NewText);
        Assert.NotEqual(moved.Old!.Start, moved.New!.Start);
        Assert.Contains(changes.Changes, static change => change.Kind == OfficeChangeKind.Inserted && change.NewText == "added words");
        Assert.Contains(changes.Changes, static change => change.Kind == OfficeChangeKind.Replaced && change.OldText == "Old" && change.NewText == "New");
        Assert.True(changes.IsComplete);
        Assert.Equal(Enumerable.Range(1, changes.Changes.Count), changes.Changes.Select(static change => change.Id));
    }

    [Fact]
    public void ChangeSetJsonIsEscapedAndComplete() {
        var regions = new[] { new OfficeChangeRegion(2, 10, 20, 30.12345, 40) };
        var set = new OfficeChangeSet(new[] {
            new OfficeChange(1, OfficeChangeKind.Replaced, "say \"hi\"\n", "tab\there", new OfficeChangeAnchor(0, 4, 1), new OfficeChangeAnchor(0, 3, 2, "p1", regions))
        }, "a.pdf", null, "word");

        string json = set.ToJson(indented: false);

        Assert.Equal(
            "{\"oldSource\":\"a.pdf\",\"newSource\":null,\"unit\":\"word\",\"isComplete\":true," +
            "\"summary\":{\"total\":1,\"inserted\":0,\"deleted\":0,\"replaced\":1,\"moved\":0}," +
            "\"changes\":[{\"id\":1,\"kind\":\"replaced\",\"oldText\":\"say \\\"hi\\\"\\n\",\"newText\":\"tab\\there\"," +
            "\"old\":{\"start\":0,\"length\":4,\"pageNumber\":1}," +
            "\"new\":{\"start\":0,\"length\":3,\"pageNumber\":2,\"location\":\"p1\",\"regions\":[{\"pageNumber\":2,\"left\":10,\"bottom\":20,\"right\":30.123,\"top\":40}]}}]}",
            json);
        Assert.Contains("\n  \"changes\": [\n", set.ToJson());
    }

    [Fact]
    public void SideBySideHtmlHighlightsWordsAndEscapes() {
        string oldText = "first line\nkeep <me>\nthe old value\n";
        string newText = "first line\nkeep <me>\nthe new value\nadded line\n";

        string html = OfficeDiffHtmlReport.SideBySide(oldText, newText, new OfficeDiffHtmlOptions { Title = "Report & co" });

        Assert.Contains("<title>Report &amp; co</title>", html);
        Assert.Contains("keep &lt;me&gt;", html);
        Assert.Contains("<del>old</del>", html);
        Assert.Contains("<ins>new</ins>", html);
        Assert.Contains("<td class=\"ins\">added line</td>", html);
        Assert.DoesNotContain("<script", html, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SideBySideHtmlCollapsesLongUnchangedStretches() {
        string oldText = string.Join("\n", Enumerable.Range(1, 30).Select(static index => "line " + index));
        string newText = oldText.Replace("line 15", "line fifteen");

        string html = OfficeDiffHtmlReport.SideBySide(oldText, newText, new OfficeDiffHtmlOptions { ContextLines = 2 });

        Assert.Contains("12 unchanged line(s)", html);
        Assert.Contains("13 unchanged line(s)", html);
        Assert.Contains("line 13", html);
        Assert.DoesNotContain(">line 11<", html);
    }

    private static void AssertCoversBothSides<T>(OfficeDiffResult<T> diff) {
        int oldCursor = 0, newCursor = 0;
        foreach (OfficeDiffEdit edit in diff.Edits) {
            Assert.Equal(oldCursor, edit.OldStart);
            Assert.Equal(newCursor, edit.NewStart);
            oldCursor += edit.OldLength;
            newCursor += edit.NewLength;
        }

        Assert.Equal(diff.Old.Count, oldCursor);
        Assert.Equal(diff.New.Count, newCursor);
    }

    private static int[] RandomSequence(Random random, int length, int alphabet) {
        int[] items = new int[length];
        for (int index = 0; index < length; index++) items[index] = random.Next(alphabet);
        return items;
    }

    private static int[] Mutate(Random random, int[] source, int alphabet) {
        var items = source.ToList();
        int edits = random.Next(0, 8);
        for (int edit = 0; edit < edits; edit++) {
            int kind = random.Next(3);
            if (kind == 0 || items.Count == 0) items.Insert(random.Next(items.Count + 1), random.Next(alphabet));
            else if (kind == 1) items.RemoveAt(random.Next(items.Count));
            else items[random.Next(items.Count)] = random.Next(alphabet);
        }

        return items.ToArray();
    }

    private static string RandomText(Random random, string[] vocabulary) {
        var builder = new StringBuilder();
        int length = random.Next(0, 30);
        for (int index = 0; index < length; index++) builder.Append(vocabulary[random.Next(vocabulary.Length)]);
        return builder.ToString();
    }

    private static string MutateText(Random random, string text, string[] vocabulary) {
        var builder = new StringBuilder(text);
        int edits = random.Next(1, 5);
        for (int edit = 0; edit < edits; edit++) {
            int position = random.Next(builder.Length + 1);
            if (random.Next(2) == 0 || builder.Length == 0) builder.Insert(position, vocabulary[random.Next(vocabulary.Length)]);
            else builder.Remove(Math.Min(position, builder.Length - 1), 1);
        }

        return builder.ToString();
    }

    private static int ReferenceLcsLength<T>(T[] a, T[] b) {
        var comparer = EqualityComparer<T>.Default;
        int[,] table = new int[a.Length + 1, b.Length + 1];
        for (int i = a.Length - 1; i >= 0; i--) {
            for (int j = b.Length - 1; j >= 0; j--) {
                table[i, j] = comparer.Equals(a[i], b[j]) ? table[i + 1, j + 1] + 1 : Math.Max(table[i + 1, j], table[i, j + 1]);
            }
        }

        return table[0, 0];
    }

    private static string Concat(IEnumerable<OfficeTextToken> tokens) {
        var builder = new StringBuilder();
        foreach (OfficeTextToken token in tokens) builder.Append(token.Text);
        return builder.ToString();
    }
}
