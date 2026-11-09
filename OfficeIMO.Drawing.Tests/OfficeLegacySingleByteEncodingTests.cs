using System.Text;
using OfficeIMO.Core.Internal;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class OfficeLegacySingleByteEncodingTests {
    [Fact]
    public void Windows1257PreservesBalticLettersAndRejectsBestFitEncoding() {
        const string expected = "ĀāĒēĢģĪī ĄąČčĘęĖėĮįŠšŲųÕõ €";
        byte[] bytes = {
            0xC2, 0xE2, 0xC7, 0xE7, 0xCC, 0xEC, 0xCE, 0xEE, 0x20,
            0xC0, 0xE0, 0xC8, 0xE8, 0xC6, 0xE6, 0xCB, 0xEB, 0xC1, 0xE1,
            0xD0, 0xF0, 0xD8, 0xF8, 0xD5, 0xF5, 0x20, 0x80
        };
        Assert.Equal(expected, OfficeLegacySingleByteEncoding.Decode(bytes, 0, bytes.Length, 1257));
        Assert.Equal(bytes, OfficeLegacySingleByteEncoding.Encode(expected, 1257));
        Assert.Throws<EncoderFallbackException>(() => OfficeLegacySingleByteEncoding.Encode("漢", 1257));
    }
}
