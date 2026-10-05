using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, 0D)]
    [InlineData(false, 1.5D)]
    [InlineData(false, 1638D)]
    [InlineData(true, 0D)]
    [InlineData(true, 1.5D)]
    [InlineData(true, 1638D)]
    public void RunKerningThresholdPreservesNativeHalfPoints(bool nativeDoc, double threshold) {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("AB").KerningMinimumFontSizePoints = threshold;
        using WordDocument reloaded = WordDocument.Load(new MemoryStream(nativeDoc
            ? source.ToBytes(WordFileFormat.Doc) : source.ToBytes()));
        Assert.Equal(threshold, reloaded.Paragraphs[0].KerningMinimumFontSizePoints);
    }

    [Fact]
    public void RunKerningThresholdCanReturnToInheritanceAndRejectsInvalidValuesAtomically() {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("AB");
        paragraph.KerningMinimumFontSizePoints = 1.25D;
        Assert.Equal(1.5D, paragraph.KerningMinimumFontSizePoints);
        foreach (double invalid in new[] { -0.5D, 1638.5D, double.NaN, double.PositiveInfinity }) {
            Assert.Throws<ArgumentOutOfRangeException>(() => paragraph.KerningMinimumFontSizePoints = invalid);
            Assert.Equal(1.5D, paragraph.KerningMinimumFontSizePoints);
        }
        paragraph.KerningMinimumFontSizePoints = null;
        using WordDocument reloaded = WordDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Null(reloaded.Paragraphs[0].KerningMinimumFontSizePoints);
    }
}
