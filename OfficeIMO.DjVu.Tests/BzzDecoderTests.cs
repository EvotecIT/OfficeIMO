using OfficeIMO.DjVu;

namespace OfficeIMO.DjVu.Tests;

public sealed class BzzDecoderTests {
    [Theory]
    [InlineData("short")]
    [InlineData("binary")]
    [InlineData("multiblock")]
    public void ReferenceEncodedBlocksRecoverOriginalBytes(string name) {
        byte[] expected = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Bzz", name + ".bin"));
        byte[] encoded = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Bzz", name + ".bzz"));
        var budget = new DjVuReadBudget(new DjVuReadOptions(), default);
        Assert.Equal(expected, BzzDecoder.Decode(encoded, 0, encoded.Length, budget));
    }

    [Fact]
    public void ExpandedLimitIsEnforcedBeforeAllocatingReferenceBlock() {
        byte[] encoded = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Bzz", "binary.bzz"));
        var budget = new DjVuReadBudget(new DjVuReadOptions { MaxExpandedBytes = 8 }, default);
        var exception = Assert.Throws<DjVuResourceLimitException>(() => BzzDecoder.Decode(encoded, 0, encoded.Length, budget));
        Assert.Equal(nameof(DjVuReadOptions.MaxExpandedBytes), exception.LimitName);
    }
}
