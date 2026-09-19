using System.Collections.Generic;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfIndirectObjectSerializationTests {
    [Fact]
    public void DirectStreamSerializationPreservesIndirectObjectBytes() {
        var sourceObjects = new Dictionary<int, PdfIndirectObject> {
            [7] = new PdfIndirectObject(7, 1, new PdfName("Example"))
        };
        var numberMap = new Dictionary<int, int> { [7] = 2 };
        var context = new PdfPageExtractor.SerializationContext(
            numberMap,
            pagesObjectId: 0,
            new Dictionary<int, Dictionary<string, PdfObject>>(),
            sourceObjects);
        var dictionary = new PdfDictionary();
        dictionary.Items["Type"] = new PdfName("XObject");
        dictionary.Items["Length"] = new PdfNumber(1);
        dictionary.Items["Related"] = new PdfReference(7, 1);
        var stream = new PdfStream(dictionary, new byte[] { 0, 10, 13, 37, 128, 255 });

        byte[] expected = PdfPageExtractor.WrapObject(4, PdfPageExtractor.SerializeObject(stream, context));
        byte[] actual = PdfPageExtractor.SerializeIndirectObject(4, stream, context);

        Assert.Equal(expected, actual);
    }
}
