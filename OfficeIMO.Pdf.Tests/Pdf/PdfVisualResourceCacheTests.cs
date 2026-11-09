using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfVisualResourceCacheTests {
    [Fact]
    public void SameDictionaryStreamsBeyondCandidateLimitAreStillReused() {
        const string dictionary = "<< /Length 4 >>";
        var cache = new PdfVisualResourceCache();
        for (int i = 0; i < 12; i++) cache.AddStream(dictionary, new byte[] { 1, 2, 3, (byte)i }, 100 + i);

        for (int i = 0; i < 12; i++) {
            Assert.True(cache.TryGetStream(dictionary, new byte[] { 1, 2, 3, (byte)i }, out int id));
            Assert.Equal(100 + i, id);
        }
        Assert.False(cache.TryGetStream(dictionary, new byte[] { 1, 2, 3, 99 }, out _));
    }
}
