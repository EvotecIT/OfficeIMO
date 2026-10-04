using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PngContainerValidationTests {
    [Fact]
    public void ContainerValidationCannotAuthorizeDifferentPngBytes() {
        byte[] original = OfficePngWriter.Encode(new OfficeRasterImage(3, 2, OfficeColor.SteelBlue));
        Assert.True(OfficePngContainerValidation.TryCreate(original, CancellationToken.None, out var validation));
        Assert.True(OfficePngReader.TryDecode(original, CancellationToken.None, 0L, out var decoded, validation));
        Assert.Equal(OfficeColor.SteelBlue, decoded!.GetPixel(2, 1));

        byte[] damaged = (byte[])original.Clone();
        damaged[29] ^= 1; // An IHDR CRC mismatch must be checked for this different input.
        Assert.False(OfficePngReader.TryDecode(damaged, CancellationToken.None, 0L, out decoded, validation));
        Assert.Null(decoded);
    }
}
