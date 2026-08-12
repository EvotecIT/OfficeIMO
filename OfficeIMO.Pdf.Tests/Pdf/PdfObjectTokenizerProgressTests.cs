using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfObjectTokenizerProgressTests {
    [Theory]
    [InlineData(")")]
    [InlineData(">")]
    public void ParseObjects_ConsumesUnexpectedStandaloneDelimiter(string delimiter) {
        byte[] pdf = Encoding.ASCII.GetBytes(
            "%PDF-1.4\n" +
            "1 0 obj\n" + delimiter + "\nendobj\n" +
            "trailer\n<< /Root 1 0 R >>\n" +
            "%%EOF\n");
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxTokensPerObject = 8 }
        };

        var (objects, _) = PdfSyntax.ParseObjects(pdf, options);

        Assert.True(objects.ContainsKey(1));
    }
}
