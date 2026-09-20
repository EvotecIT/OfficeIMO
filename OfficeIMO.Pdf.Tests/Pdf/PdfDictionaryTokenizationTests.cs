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

    [Fact]
    public void DictionaryRangeUsesItsOwnCharacterBudgetAndDoesNotReadTheFollowingObject() {
        PdfDictionary dictionary = ParseDictionary(
            "/Payload (range-safe) /Nested << /Values [1 2 3] >> /Hex <4142>",
            new PdfLoadOptions { Limits = new PdfReadLimits { MaxObjectCharacters = 128 } });

        Assert.Equal(3, dictionary.Items.Count);
        Assert.Equal("range-safe", Assert.IsType<PdfStringObj>(dictionary.Items["Payload"]).Value);
        PdfDictionary nested = Assert.IsType<PdfDictionary>(dictionary.Items["Nested"]);
        Assert.Equal(3, Assert.IsType<PdfArray>(nested.Items["Values"]).Items.Count);
        Assert.Equal("AB", Assert.IsType<PdfStringObj>(dictionary.Items["Hex"]).Value);
        Assert.False(dictionary.HasIncompleteSyntax);
    }

    [Fact]
    public void MalformedDelimiterAtEndOfDictionaryRangeDoesNotConsumeFollowingObject() {
        PdfDictionary dictionary = ParseDictionary("/Good 3 )");

        Assert.Single(dictionary.Items);
        Assert.Equal(3, Assert.IsType<PdfNumber>(dictionary.Items["Good"]).Value);
        Assert.True(dictionary.HasIncompleteSyntax);
    }

    [Fact]
    public void CommonAndEscapedNamesKeepEquivalentDecodedKeysAndValues() {
        PdfDictionary dictionary = ParseDictionary("/Type /Page /Ty#70eAlias /Pa#67e /Custom#20Key /Value#23Name");

        Assert.Equal("Page", Assert.IsType<PdfName>(dictionary.Items["Type"]).Name);
        Assert.Equal("Page", Assert.IsType<PdfName>(dictionary.Items["TypeAlias"]).Name);
        Assert.Equal("Value#Name", Assert.IsType<PdfName>(dictionary.Items["Custom Key"]).Name);
        Assert.False(dictionary.HasIncompleteSyntax);
    }

    [Fact]
    public void DuplicateKeysKeepTheLastValueAndMarkDictionariesIncomplete() {
        PdfDictionary dictionary = ParseDictionary("/Value 1 /Value 2 /Nested << /Value 3 /Value 4 >>");

        Assert.Equal(2, Assert.IsType<PdfNumber>(dictionary.Items["Value"]).Value);
        PdfDictionary nested = Assert.IsType<PdfDictionary>(dictionary.Items["Nested"]);
        Assert.Equal(4, Assert.IsType<PdfNumber>(nested.Items["Value"]).Value);
        Assert.True(dictionary.HasIncompleteSyntax);
        Assert.True(nested.HasIncompleteSyntax);
    }

    private static PdfDictionary ParseDictionary(string entries, PdfLoadOptions? options = null) {
        byte[] pdf = Encoding.ASCII.GetBytes(
            "%PDF-1.7\n" +
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n" +
            "2 0 obj\n<< /Type /Pages /Count 0 /Kids [] >>\nendobj\n" +
            "3 0 obj\n<< " + entries + " >>\nendobj\n" +
            "trailer\n<< /Root 1 0 R /Size 4 >>\nstartxref\n0\n%%EOF\n");
        return Assert.IsType<PdfDictionary>(PdfSyntax.ParseObjects(pdf, options).Map[3].Value);
    }
}
