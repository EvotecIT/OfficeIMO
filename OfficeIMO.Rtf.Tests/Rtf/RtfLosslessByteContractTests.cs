using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public sealed class RtfLosslessByteContractTests {
    [Fact]
    public void CharacterInputReportsWhenNoByteLosslessRepresentationExists() {
        const string source = "{\\rtf1 Unicode ż}";

        RtfReadResult result = RtfDocument.Read(source);

        Assert.False(result.HasOriginalBytes);
        Assert.False(result.CanWriteLosslessBytes);
        Assert.False(result.TryGetLosslessBytes(out byte[] bytes));
        Assert.Empty(bytes);
        Assert.Throws<InvalidOperationException>(() => result.ToBytesLossless());
        Assert.Equal(source, result.ToRtfLossless());
    }

    [Fact]
    public void ByteInputReturnsAnImmutableCopyOfTheExactSourceBytes() {
        byte[] source = { (byte)'{', (byte)'\\', (byte)'r', (byte)'t', (byte)'f', (byte)'1', (byte)' ', 0xFF, (byte)'}' };

        RtfReadResult result = RtfDocument.LoadResult(source);
        source[0] = 0;
        byte[] first = result.ToBytesLossless();
        first[0] = 0;

        Assert.True(result.HasOriginalBytes);
        Assert.True(result.CanWriteLosslessBytes);
        Assert.Equal((byte)'{', result.ToBytesLossless()[0]);
        Assert.Equal(0xFF, result.ToBytesLossless()[7]);
    }

    [Fact]
    public void ExplicitlyDecodedStreamStillRetainsOriginalBytes() {
        const string source = "{\\rtf1 Unicode ż}";
        byte[] sourceBytes = Encoding.UTF8.GetBytes(source);
        using var stream = new MemoryStream(sourceBytes);

        RtfReadResult result = RtfDocument.LoadResult(stream, encoding: Encoding.UTF8);

        Assert.True(result.HasOriginalBytes);
        Assert.True(result.TryGetLosslessBytes(out byte[] bytes));
        Assert.Equal(sourceBytes, bytes);
    }

    [Fact]
    public async Task AsyncLosslessSavePreservesExactOriginalBytesIncludingUtf8Bom() {
        byte[] payload = Encoding.UTF8.GetBytes("{\\rtf1 Unicode ż}");
        byte[] sourceBytes = Encoding.UTF8.GetPreamble().Concat(payload).ToArray();
        using var input = new MemoryStream(sourceBytes, writable: false);
        RtfReadResult result = RtfDocument.LoadResult(input, encoding: Encoding.UTF8);
        using var stream = new MemoryStream();

        await result.SaveLosslessAsync(stream);

        Assert.Equal(sourceBytes, stream.ToArray());
    }
}
