using OfficeIMO.Core.Internal;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingZlibChecksumTests {
    [Theory]
    [InlineData(31)]
    [InlineData(32)]
    [InlineData(33)]
    [InlineData(5551)]
    [InlineData(5552)]
    [InlineData(5553)]
    [InlineData(65537)]
    public void Adler32PreservesWeightedByteOrderAndModuloBoundaries(int length) {
        foreach (bool maximumBytes in new[] { false, true }) {
            var bytes = new byte[length];
            uint a = 1;
            uint b = 0;
            for (int index = 0; index < length; index++) {
                bytes[index] = maximumBytes ? (byte)255 : unchecked((byte)(index * 73 + 19));
                // Independent RFC recurrence reduces after every byte, rather
                // than sharing the production block accumulation or SIMD helpers.
                a = (a + bytes[index]) % 65521;
                b = (b + a) % 65521;
            }
            Assert.Equal((b << 16) | a, OfficeZlibCodec.Adler32(bytes));
        }
    }
}
