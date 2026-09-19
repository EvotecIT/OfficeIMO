using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfParsedStringOwnershipTests {
    [Theory]
    [InlineData("(A\\050B)", "A(B", new byte[] { 65, 40, 66 })]
    [InlineData("<414228>", "AB(", new byte[] { 65, 66, 40 })]
    public void ParsedStringsPreserveDecodedAndEncodedValues(string token, string expectedValue, byte[] expectedBytes) {
        byte[] pdf = Encoding.ASCII.GetBytes($"%PDF-1.7\n1 0 obj\n{token}\nendobj\n%%EOF\n");

        PdfStringObj parsed = Assert.IsType<PdfStringObj>(Assert.Single(PdfSyntax.ParseObjects(pdf).Map).Value.Value);

        Assert.Equal(expectedValue, parsed.Value);
        Assert.Equal(expectedBytes, parsed.RawBytes);
        Assert.Equal(token.Length, parsed.EncodedTokenLength);
    }

    [Fact]
    public void ByteArrayConstructorStillCopiesCallerOwnedBytes() {
        byte[] input = { 65, 66 };
        var value = new PdfStringObj(input);

        input[0] = 90;

        Assert.Equal("AB", value.Value);
        Assert.Equal(new byte[] { 65, 66 }, value.RawBytes);
    }
}
