using System.Text;
using OfficeIMO.TextAnalysis;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>The full Unicode 16.0.0 UAX #29 conformance files for word and grapheme cluster boundaries.</summary>
public class TextUnicodeSegmentationConformanceTests {
    private static string Fixture(string name) => Path.Combine(AppContext.BaseDirectory, "TestAssets", "Unicode", name);

    [Fact]
    public void WordBoundariesMatchWordBreakTest() {
        int cases = RunConformance("WordBreakTest.txt", UnicodeTextSegmenter.GetWordBoundaries);
        Assert.Equal(1826, cases);
    }

    [Fact]
    public void GraphemeClusterBoundariesMatchGraphemeBreakTest() {
        int cases = RunConformance("GraphemeBreakTest.txt", UnicodeTextSegmenter.GetGraphemeClusterBoundaries);
        Assert.Equal(1093, cases);
    }

    [Fact]
    public void TablesComeFromThePinnedUnicodeVersion() {
        Assert.Equal("16.0.0", UnicodeTextSegmenter.UnicodeVersion);
        Assert.StartsWith("# WordBreakTest-16.0.0.txt", File.ReadLines(Fixture("WordBreakTest.txt")).First());
        Assert.StartsWith("# GraphemeBreakTest-16.0.0.txt", File.ReadLines(Fixture("GraphemeBreakTest.txt")).First());
    }

    [Theory]
    [InlineData("Hello, world!", new[] { "Hello", "world" })]
    [InlineData("can't 3.14 e-mail", new[] { "can't", "3.14", "e", "mail" })]
    [InlineData("你好 Zażółć", new[] { "你", "好", "Zażółć" })]
    [InlineData("👍🏽 ok", new[] { "ok" })]
    public void WordLikeSegmentsSkipSpacesPunctuationAndEmoji(string text, string[] expected) {
        string[] words = UnicodeTextSegmenter.GetWordSegments(text)
            .Where(static segment => segment.IsWordLike)
            .Select(segment => segment.GetText(text))
            .ToArray();
        Assert.Equal(expected, words);
    }

    [Fact]
    public void GraphemeClustersKeepEmojiSequencesAndCombiningMarksTogether() {
        string text = "é👩‍💻🇵🇱";
        string[] clusters = UnicodeTextSegmenter.GetGraphemeClusters(text).Select(cluster => cluster.GetText(text)).ToArray();
        Assert.Equal(new[] { "é", "👩‍💻", "🇵🇱" }, clusters);
    }

    [Fact]
    public void EmptyTextHasNoBoundaries() {
        Assert.Empty(UnicodeTextSegmenter.GetWordBoundaries(string.Empty));
        Assert.Empty(UnicodeTextSegmenter.GetGraphemeClusters(string.Empty));
    }

    private static int RunConformance(string name, Func<string, IReadOnlyList<int>> boundaries) {
        int cases = 0;
        foreach (string raw in File.ReadLines(Fixture(name))) {
            string line = raw.Split('#')[0].Trim();
            if (line.Length == 0) continue;
            var text = new StringBuilder();
            var expected = new List<int>();
            foreach (string token in line.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)) {
                if (token == "÷") expected.Add(text.Length);
                else if (token != "×") {
                    int codePoint = Convert.ToInt32(token, 16);
                    text.Append(codePoint >= 0x10000 ? char.ConvertFromUtf32(codePoint) : ((char)codePoint).ToString());
                }
            }
            Assert.True(expected.SequenceEqual(boundaries(text.ToString())), raw);
            cases++;
        }
        return cases;
    }
}
