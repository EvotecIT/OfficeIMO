using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfDictionaryTokenizationTests {
    [Fact]
    public void SparseLongStringDictionaryKeepsItsCompleteValue() {
        string payload = new string('A', 65536);
        PdfDictionary dictionary = ParseDictionary("/Payload (" + payload + ")");

        PdfStringObj value = Assert.IsType<PdfStringObj>(dictionary.Items["Payload"]);
        Assert.Equal(payload.Length, value.RawBytes.Length);
        Assert.Equal(payload, value.Value);
        Assert.False(dictionary.HasIncompleteSyntax);
    }

    [Fact]
    public void DenseDictionaryCanGrowPastItsInitialTokenCapacity() {
        var body = new StringBuilder();
        for (int i = 0; i < 300; i++) {
            body.Append("/K").Append(i).Append(' ').Append(i).Append(' ');
        }

        PdfDictionary dictionary = ParseDictionary(body.ToString());

        Assert.Equal(300, dictionary.Items.Count);
        Assert.Equal(299, Assert.IsType<PdfNumber>(dictionary.Items["K299"]).Value);
        Assert.False(dictionary.HasIncompleteSyntax);
    }

    private static PdfDictionary ParseDictionary(string entries) {
        byte[] pdf = Encoding.ASCII.GetBytes(
            "%PDF-1.7\n" +
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n" +
            "2 0 obj\n<< /Type /Pages /Count 0 /Kids [] >>\nendobj\n" +
            "3 0 obj\n<< " + entries + " >>\nendobj\n" +
            "trailer\n<< /Root 1 0 R /Size 4 >>\nstartxref\n0\n%%EOF\n");
        return Assert.IsType<PdfDictionary>(PdfSyntax.ParseObjects(pdf).Map[3].Value);
    }
}
