using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PdfLoadPolicyCloneTests {
    [Theory]
    [InlineData(2000, 1000)]
    [InlineData(500, 500)]
    public void PerReadCloneCanTightenButNeverRelaxInputAdmission(long requested, long expected) {
        var source = new PdfLoadOptions {
            Password = "original", ParsingMode = PdfParsingMode.Strict,
            Limits = new PdfReadLimits { MaxInputBytes = 1000, MaxPages = 7 }, IncludeImages = false
        };
        var copy = source.Clone(requested, "per-read");
        Assert.Equal(expected, copy.Limits.MaxInputBytes);
        Assert.Equal(7, copy.Limits.MaxPages);
        Assert.Equal("per-read", copy.Password);
        Assert.Equal("original", source.Password);
        Assert.Equal(1000, source.Limits.MaxInputBytes);
        Assert.Equal(PdfParsingMode.Strict, copy.ParsingMode);
        Assert.False(copy.IncludeImages);
        Assert.NotSame(source.Limits, copy.Limits);
        Assert.Equal("original", source.Clone(null, null).Password);
    }
}
