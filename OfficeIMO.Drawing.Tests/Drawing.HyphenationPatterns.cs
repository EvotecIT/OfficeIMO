using OfficeIMO.Drawing;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class DrawingHyphenationPatternTests {
    [Theory]
    [InlineData("en-US", "representation", new[] { 3, 5, 8, 10 })]
    [InlineData("en-US", "extraordinary", new[] { 2, 5, 7, 9 })]
    [InlineData("en-US", "internationalization", new[] { 2, 5, 7, 11, 13, 16 })]
    [InlineData("de-DE", "Silbentrennung", new[] { 3, 6, 10 })]
    [InlineData("de-DE", "Donaudampfschifffahrt", new[] { 2, 5, 10, 16 })]
    [InlineData("de-DE", "außergewöhnlich", new[] { 2, 5, 7, 11 })]
    [InlineData("en", "ASSOCIATE", new[] { 2, 4 })]
    [InlineData("en", "obligatory", new[] { 5, 6 })]
    [InlineData("en", "present", new int[0])]
    [InlineData("en", "projects", new int[0])]
    [InlineData("de-1996", "BA\u0308CKEREI", new[] { 3, 6 })]
    [InlineData("en-us", "(representation),", new[] { 4, 6, 9, 11 })]
    public void EmbeddedPatternsMatchIndependentPatternOracle(string language, string word, int[] expected) {
        Assert.Equal(expected, OfficeTextHyphenationPatterns.GetBreakpoints(word, language));
    }

    [Theory]
    [InlineData(null, "representation")]
    [InlineData("zz", "representation")]
    [InlineData("en-GB", "representation")]
    [InlineData("de-1901", "Silbentrennung")]
    [InlineData("en-US", "representation2")]
    [InlineData("en-US", "representation implementation")]
    [InlineData("en-US", "representation-implementation")]
    public void UnsupportedTagsAndNonWordTokensHaveNoInventedBreaks(string? language, string word) {
        Assert.Empty(OfficeTextHyphenationPatterns.GetBreakpoints(word, language));
    }

    [Fact]
    public void BoundedLookupDoesNotExposeSharedMutableResults() {
        Assert.Empty(OfficeTextHyphenationPatterns.GetBreakpoints(new string('a', 513), "en-US"));
        int[] first = Assert.IsType<int[]>(OfficeTextHyphenationPatterns.GetBreakpoints("representation", "en-US"));
        first[0] = 99;
        Parallel.For(0, 32, _ => Assert.Equal(new[] { 3, 5, 8, 10 },
            OfficeTextHyphenationPatterns.GetBreakpoints("representation", "en-US")));
    }
}
