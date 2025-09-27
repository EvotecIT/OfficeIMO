using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public class PdfSyntaxNestedDictionaryTests {
    [Fact]
    public void ParseObjectsHandlesNestedDictionaries() {
        var pdf = """
%PDF-1.4
1 0 obj
<< /Type /Catalog /Pages 2 0 R >>
endobj
2 0 obj
<< /Type /Pages /Kids [3 0 R] /Count 1 >>
endobj
3 0 obj
<< /Type /Page
   /Parent 2 0 R
   /Resources <<
       /Font <<
           /F1 <<
               /Type /Font
               /Subtype /Type1
               /BaseFont /Helvetica
               /Encoding << /Type /Encoding /BaseEncoding /WinAnsiEncoding >>
           >>
       >>
       /ExtGState <<
           /GS1 <<
               /Type /ExtGState
               /SMask <<
                   /Type /Mask
                   /S /Luminosity
                   /G 5 0 R
               >>
           >>
       >>
   >>
   /MediaBox [0 0 612 792]
   /Contents 4 0 R
>>
endobj
4 0 obj
<< /Type /XObject
   /Subtype /Form
   /Resources <<
       /Pattern <<
           /P1 <<
               /Type /Pattern
               /PaintType 1
               /TilingType 1
           >>
       >>
   >>
   /Length 4
>>
stream
q
Q
endstream
endobj
5 0 obj
<< /Type /XObject /Subtype /Image >>
endobj
trailer
<< /Root 1 0 R /Size 6 >>
%%EOF
""";

        var bytes = PdfEncoding.Latin1GetBytes(pdf);
        var (map, trailerRaw) = PdfSyntax.ParseObjects(bytes);

        Assert.True(map.TryGetValue(3, out var pageObject));
        var pageDictionary = Assert.IsType<PdfDictionary>(pageObject.Value);
        var resources = Assert.IsType<PdfDictionary>(pageDictionary.Items["Resources"]);

        var fonts = Assert.IsType<PdfDictionary>(resources.Items["Font"]);
        var f1 = Assert.IsType<PdfDictionary>(fonts.Items["F1"]);
        var encoding = Assert.IsType<PdfDictionary>(f1.Items["Encoding"]);
        Assert.Equal("Encoding", Assert.IsType<PdfName>(encoding.Items["Type"]).Name);
        Assert.Equal("WinAnsiEncoding", Assert.IsType<PdfName>(encoding.Items["BaseEncoding"]).Name);

        Assert.True(resources.Items.ContainsKey("ExtGState"), string.Join(", ", resources.Items.Keys));
        var extGState = Assert.IsType<PdfDictionary>(resources.Items["ExtGState"]);
        var gs1 = Assert.IsType<PdfDictionary>(extGState.Items["GS1"]);
        var sMask = Assert.IsType<PdfDictionary>(gs1.Items["SMask"]);
        Assert.Equal("Mask", Assert.IsType<PdfName>(sMask.Items["Type"]).Name);
        var gReference = Assert.IsType<PdfReference>(sMask.Items["G"]);
        Assert.Equal(5, gReference.ObjectNumber);

        Assert.True(map.TryGetValue(4, out var streamObject));
        var stream = Assert.IsType<PdfStream>(streamObject.Value);
        var streamResources = Assert.IsType<PdfDictionary>(stream.Dictionary.Items["Resources"]);
        var patternDict = Assert.IsType<PdfDictionary>(streamResources.Items["Pattern"]);
        var p1 = Assert.IsType<PdfDictionary>(patternDict.Items["P1"]);
        Assert.Equal(1d, Assert.IsType<PdfNumber>(p1.Items["PaintType"]).Value);
        Assert.Equal("Pattern", Assert.IsType<PdfName>(p1.Items["Type"]).Name);

        var streamContent = PdfEncoding.Latin1GetString(stream.Data);
        Assert.Equal("q\nQ", streamContent);

        Assert.Contains("trailer", trailerRaw, StringComparison.OrdinalIgnoreCase);
    }
}
