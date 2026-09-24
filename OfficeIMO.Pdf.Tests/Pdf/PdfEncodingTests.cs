using System.Linq;
using System.Text;
using System.Threading;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfEncodingTests {
    [Fact]
    public void Latin1DecodingPreservesEveryByteIncludingBinaryPdfContent() {
        byte[] bytes = Enumerable.Range(0, 256).Select(static value => (byte)value).ToArray();
        string expected = new string(bytes.Select(static value => (char)value).ToArray());

        Assert.Equal(expected, PdfEncoding.Latin1GetString(bytes));
        Assert.Equal(expected.Substring(31, 193), PdfEncoding.Latin1GetString(bytes, 31, 193));
        Assert.Equal(bytes, PdfEncoding.Latin1GetBytes(PdfEncoding.Latin1GetString(bytes)));
    }

    [Fact]
    public void CancellableStringMaterializationPreservesLargeLegacyTargetSlices() {
        string source = new string('a', 70_000) + "\u03A9" + new string('b', 70_000);
        char[] characters = source.ToCharArray();
        var builder = new StringBuilder("prefix").Append(source).Append("suffix");
        using var active = new CancellationTokenSource();
        CancellationToken token = active.Token;

        Assert.Equal(source, PdfEncoding.CharArrayToStringCancellable(characters, token));
        Assert.Equal(source, PdfEncoding.StringSliceCancellable("prefix" + source + "suffix", 6, source.Length, token));
        Assert.Equal(source, PdfEncoding.StringBuilderToStringCancellable(builder, 6, source.Length, token));

        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() => PdfEncoding.CharArrayToStringCancellable(characters, cancelled.Token));
    }
}
