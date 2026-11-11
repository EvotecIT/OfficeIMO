using System.Text;
using OfficeIMO.Core.Internal;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class Base64CodecTests {
    [Theory]
    [InlineData("AB==", 1)]
    [InlineData("AAB=", 2)]
    [InlineData("AB", 1)]
    [InlineData("AAB", 2)]
    public void RetainedLengthCountsTheWholeRawRangeWhenTheByteDecoderRejectsItsFinalQuantum(
        string tail, int terminalBytes) {
        const int quanta = 4096;
        string payload = "\r\n" + new string('A', quanta * 4) + " \t" + tail + "\v\f";
        byte[] input = Encoding.ASCII.GetBytes("prefix:" + payload + ":suffix");
        bool retainsRaw = false;
#if NET8_0_OR_GREATER
        var target = System.Reflection.CustomAttributeExtensions.GetCustomAttribute<
            System.Runtime.Versioning.TargetFrameworkAttribute>(typeof(OfficeBase64).Assembly)!;
        if (target.FrameworkName.StartsWith(".NETCoreApp", StringComparison.Ordinal)) {
            // The platform byte decoder's policy differs from Convert and from the scalar codec.
            byte[] paddedTail = Encoding.ASCII.GetBytes(tail.PadRight(4, '='));
            retainsRaw = System.Buffers.Text.Base64.DecodeFromUtf8(paddedTail, new byte[3],
                out _, out _) != System.Buffers.OperationStatus.Done;
        }
#endif
        long expected = retainsRaw ? payload.Length : quanta * 3 + terminalBytes;

        long charged = OfficeBase64.GetMimeDecodedLength(input, 7, payload.Length,
            out bool valid, out bool recovered);

        Assert.Equal(expected, charged);
        Assert.Equal(!retainsRaw, valid);
        Assert.Equal(!retainsRaw && tail.IndexOf('=') < 0, recovered);
    }

    [Theory]
    [InlineData("AQIDBA==")]
    [InlineData(" A\tQ\rI\nDBA= = ")]
    public void PlainStringSlicesKeepConvertWhitespacePaddingAndPayloadRules(string value) {
        Assert.Equal(new byte[] { 1, 2, 3, 4 }, OfficeBase64.FromBase64String("prefix:" + value, 7));
        Assert.Empty(OfficeBase64.FromBase64String("prefix: \t\r\n", 7));
    }

    [Theory]
    [InlineData("AA")]
    [InlineData("AA=")]
    [InlineData("AA==AAAA")]
    [InlineData("AA\v==")]
    [InlineData("AA\f==")]
    [InlineData("AA\u00a0==")]
    public void PlainStringSlicesRejectTheSameMalformedInputAsConvert(string value) {
        var expected = Assert.Throws<FormatException>(() => Convert.FromBase64String(value));
        var actual = Assert.Throws<FormatException>(() => OfficeBase64.FromBase64String("prefix:" + value, 7));
        Assert.Equal(expected.Message, actual.Message);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(15)]
    [InlineData(16)]
    [InlineData(17)]
    [InlineData(31)]
    [InlineData(32)]
    [InlineData(33)]
    [InlineData(57)]
    [InlineData(65535)]
    [InlineData(65536)]
    [InlineData(65537)]
    public void ByteKernelsAndChunkedMimeAnalysisPreservePayloadAndBufferBounds(int length) {
        var payload = new byte[length + 2];
        new Random(64).NextBytes(payload);
        string expected = Convert.ToBase64String(payload, 1, length);
        var encoded = Enumerable.Repeat((byte)0xCC, expected.Length + 4).ToArray();
        int written = OfficeBase64.EncodeToUtf8(payload, 1, length, encoded, 2);
        Assert.Equal(expected, Encoding.ASCII.GetString(encoded, 2, written));
        Assert.Equal(0xCC, encoded[0]);
        Assert.Equal(0xCC, encoded[1]);
        Assert.Equal(0xCC, encoded[written + 2]);
        Assert.Equal(0xCC, encoded[written + 3]);

        // Whitespace can interrupt a quartet, padding or a chunk, including MIME-only FF/VT.
        string mime = " " + string.Join("\r\n\t\v\f", Enumerable.Range(0, (expected.Length + 18) / 19)
            .Select(index => expected.Substring(index * 19, Math.Min(19, expected.Length - index * 19))));
        foreach (bool omitPadding in new[] { false, true }) {
            byte[] input = Encoding.ASCII.GetBytes(omitPadding ? mime.Replace("=", "") : mime);
            var analysis = new OfficeBase64.MimeAnalysis();
            for (int offset = 0; offset < input.Length; offset += 31)
                analysis.Append(input, offset, Math.Min(31, input.Length - offset));
            Assert.True(analysis.IsValid);
            Assert.Equal(length, analysis.DecodedLength);
            Assert.Equal(omitPadding && length % 3 != 0, analysis.RecoveredPadding);
            var decoded = new byte[length];
            Assert.True(OfficeBase64.TryDecodeMime(input, 0, input.Length, decoded, out written, out bool recovered));
            Assert.Equal(length, written);
            Assert.Equal(analysis.RecoveredPadding, recovered);
            Assert.Equal(payload.Skip(1).Take(length), decoded);
        }
    }

    [Theory]
    [InlineData("A")]
    [InlineData("AA=")]
    [InlineData("AAA==")]
    [InlineData("A===")]
    [InlineData("====")]
    [InlineData("AA==AAAA")]
    [InlineData("AA==AA==")]
    [InlineData("A!AA")]
    [InlineData("AA\0AA")]
    [InlineData("AA\u0080AA")]
    public void InvalidMimeIsRejectedAcrossVectorAndChunkBoundaries(string invalid) {
        foreach (int prefix in new[] { 0, 16, 32, 64, 65536 }) {
            byte[] input = (new string('A', prefix) + invalid).Select(character => (byte)character).ToArray();
            Assert.Equal(input.Length, OfficeBase64.GetMimeDecodedLength(input, 0, input.Length,
                out bool valid, out bool recovered));
            Assert.False(valid);
            Assert.False(recovered);
            Assert.False(OfficeBase64.TryDecodeMime(input, 0, input.Length, new byte[input.Length], out _, out _));
            var analysis = new OfficeBase64.MimeAnalysis();
            for (int index = 0; index < input.Length; index++) analysis.Append(input, index, 1);
            Assert.False(analysis.IsValid);
        }
    }
}
