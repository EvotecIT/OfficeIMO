using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(1, WordParagraphStyles.Heading1)]
    [InlineData(2, WordParagraphStyles.Heading2)]
    [InlineData(3, WordParagraphStyles.Heading3)]
    [InlineData(4, WordParagraphStyles.Heading4)]
    [InlineData(5, WordParagraphStyles.Heading5)]
    [InlineData(6, WordParagraphStyles.Heading6)]
    [InlineData(7, WordParagraphStyles.Heading7)]
    [InlineData(8, WordParagraphStyles.Heading8)]
    [InlineData(9, WordParagraphStyles.Heading9)]
    [InlineData(0, WordParagraphStyles.Heading1)]
    public void Test_GetHeadingStyleForLevel(int level, WordParagraphStyles expected) {
        var style = HeadingStyleMapper.GetHeadingStyleForLevel(level);
        Assert.Equal(expected, style);
    }

    [Theory]
    [InlineData(WordParagraphStyles.Heading1, 1)]
    [InlineData(WordParagraphStyles.Heading2, 2)]
    [InlineData(WordParagraphStyles.Heading3, 3)]
    [InlineData(WordParagraphStyles.Heading4, 4)]
    [InlineData(WordParagraphStyles.Heading5, 5)]
    [InlineData(WordParagraphStyles.Heading6, 6)]
    [InlineData(WordParagraphStyles.Heading7, 7)]
    [InlineData(WordParagraphStyles.Heading8, 8)]
    [InlineData(WordParagraphStyles.Heading9, 9)]
    [InlineData(WordParagraphStyles.Normal, 0)]
    public void Test_GetLevelForHeadingStyle(WordParagraphStyles style, int expected) {
        var level = HeadingStyleMapper.GetLevelForHeadingStyle(style);
        Assert.Equal(expected, level);
    }
}
