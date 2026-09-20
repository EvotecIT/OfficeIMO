using System.Linq;
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
}
