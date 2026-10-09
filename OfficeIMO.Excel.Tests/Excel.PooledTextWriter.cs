using System.Text;
using OfficeIMO.SharedSource.IO;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public class PooledTextWriterTests {
    [Fact]
    public void BufferedAsciiPrefixBeforeUnicodePreservesPreviouslyEncodedBytes() {
        using var output = new MemoryStream();
        string encodedPrefix = new string('p', 4096);
        string bufferedPrefix = new string('a', 16000);
        using (var writer = new PooledUtf8TextWriter(output, new UTF8Encoding(false), 65536, leaveOpen: true)) {
            writer.Write(encodedPrefix);
            // Character writes keep this prefix buffered even above the direct-encoding threshold.
            foreach (char character in bufferedPrefix) writer.Write(character);
            writer.WriteFragments("漢", "\ud83d", "\ude80λ");
            writer.Flush();
            Assert.Equal(Encoding.UTF8.GetBytes(encodedPrefix + bufferedPrefix + "漢🚀λ"), output.ToArray());
            writer.WriteFragments("final ", "λ", "\ud83d");
        }
        Assert.Equal(Encoding.UTF8.GetBytes(encodedPrefix + bufferedPrefix + "漢🚀λfinal λ\ufffd"), output.ToArray());
        Assert.True(output.CanWrite);
    }

    [Theory]
    [InlineData(16)]
    [InlineData(256)]
    [InlineData(4096)]
    [InlineData(65536)]
    public void FragmentWritesPreserveEncodingAcrossBufferBoundaries(int bufferSize) {
        foreach (Encoding encoding in new Encoding[] {
            new UTF8Encoding(false), new UnicodeEncoding(false, false), new UTF32Encoding(false, false)
        }) {
            using var output = new MemoryStream();
            var expected = new StringBuilder();
            using (var writer = new PooledUtf8TextWriter(output, encoding, bufferSize, leaveOpen: true)) {
                foreach (int length in new[] { 0, 1, 7, 14, 15, 16, 17, 255, 256, 257, 4095, 4096, 4097, 16383, 16384, 16385 }) {
                    string prefix = new string('a', length);
                    writer.Write(prefix);
                    writer.WriteFragments("<c r=\"A", "1048576", "\">");
                    expected.Append(prefix).Append("<c r=\"A1048576\">");
                    writer.WriteFragments("Zażółć \ud83d", "\ude80", " 漢");
                    expected.Append("Zażółć \ud83d\ude80 漢");
                    writer.Flush();
                }
                // A long fragment must retain the existing direct-encoding path.
                string longText = new string('z', 4096) + "\ud83d";
                writer.WriteFragments(longText, "\ude80", " suffix");
                expected.Append(longText).Append("\ude80 suffix");
            }
            Assert.Equal(encoding.GetBytes(expected.ToString()), output.ToArray());
            Assert.True(output.CanWrite);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FragmentWritesPreservePendingSurrogatesAndReplacement(bool longPrefix) {
        string prefix = longPrefix ? new string('a', 4096) : "prefix ";
        using var output = new MemoryStream();
        using (var writer = new PooledUtf8TextWriter(output, new UTF8Encoding(false), 16, leaveOpen: true)) {
            writer.Write(prefix + "\ud83d");
            writer.Flush();
            writer.WriteFragments("\ude80", " ASCII", "");
            writer.WriteFragments("\ud83d", "", "");
            writer.Flush();
            writer.WriteFragments("ASCII", " λ", "");
            writer.WriteFragments("\ud83d", "", "");
        }
        Assert.Equal(Encoding.UTF8.GetBytes(prefix + "🚀 ASCII\ufffdASCII λ\ufffd"), output.ToArray());
    }

    [Theory]
    [InlineData(16)]
    [InlineData(256)]
    [InlineData(4096)]
    [InlineData(65536)]
    public void MixedWritesPreserveEncodingAndOrder(int bufferSize) {
        foreach (Encoding encoding in new Encoding[] {
            new UTF8Encoding(false), new UnicodeEncoding(false, false), new UTF32Encoding(false, false)
        }) {
            using var output = new MemoryStream();
            var expected = new StringBuilder();
            using (var writer = new PooledUtf8TextWriter(output, encoding, bufferSize, leaveOpen: true)) {
                foreach (int length in new[] { 1, 15, 255, 256, 257, 4090, 4091, 4092, 4095, 4096, 4097, 65535, 65536, 65537 }) {
                    string text = new string('a', length) + "\ud83d\ude80Ł漢";
                    writer.Write('<');
                    writer.Write(text);
                    expected.Append('<').Append(text);
                    char[] characters = ("before" + text + "after").ToCharArray();
                    writer.Write(characters, 6, text.Length);
                    expected.Append(text);
#if NET6_0_OR_GREATER
                    writer.Write(text.AsSpan());
                    expected.Append(text);
#endif
                    writer.Write('\ud83d');
                    writer.Flush();
                    writer.Write('\ude80');
                    expected.Append("\ud83d\ude80");
                }
                writer.Write('\ud83d');
                expected.Append('\ud83d');
            }
            Assert.Equal(encoding.GetBytes(expected.ToString()), output.ToArray());
            Assert.True(output.CanWrite);
        }
    }

    [Fact]
    public void LongFallbackSequencesCompleteBeforeFlushReturns() {
        Encoding encoding = Encoding.GetEncoding("us-ascii",
            new EncoderReplacementFallback(new string('?', 300)), DecoderFallback.ExceptionFallback);
        string text = new string('a', 4095) + "漢";
        using var output = new MemoryStream();
        using var writer = new PooledUtf8TextWriter(output, encoding, 16, leaveOpen: true);
        writer.Write(text);
        writer.Flush();
        Assert.Equal(encoding.GetBytes(text), output.ToArray());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void EncodingFailureOnDisposeRespectsStreamOwnership(bool leaveOpen) {
        using var output = new MemoryStream();
        var writer = new PooledUtf8TextWriter(output, new UTF8Encoding(false, true), 16, leaveOpen);
        writer.Write('\ud83d');
        Assert.Throws<EncoderFallbackException>(() => writer.Dispose());
        Assert.Equal(leaveOpen, output.CanWrite);
        Assert.Throws<ObjectDisposedException>(() => writer.Write('x'));
        writer.Dispose();
    }
}
