using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfImageReferenceBudgetTests {
    [Theory]
    [InlineData("Filter")]
    [InlineData("DecodeParms")]
    public void ImageFilterReferencesAreBoundedBeforeImageDecoding(string key) {
        var dictionary = new PdfDictionary();
        dictionary.Items["Type"] = new PdfName("XObject");
        dictionary.Items["Subtype"] = new PdfName("Image");
        dictionary.Items["Width"] = new PdfNumber(1);
        dictionary.Items["Height"] = new PdfNumber(1);
        dictionary.Items["BitsPerComponent"] = new PdfNumber(8);
        dictionary.Items["ColorSpace"] = new PdfName("DeviceGray");
        dictionary.Items["Filter"] = new PdfName("FlateDecode");
        dictionary.Items[key] = new PdfReference(1, 0);
        var objects = new Dictionary<int, PdfIndirectObject> {
            [1] = new(1, 0, new PdfReference(2, 0)),
            [2] = new(2, 0, new PdfReference(3, 0)),
            [3] = new(3, 0, key == "Filter" ? new PdfName("FlateDecode") : new PdfDictionary())
        };

        Assert.Throws<InvalidDataException>(() => ResourceResolver.BuildExtractedImage(
            1, "Im1", 4, 0, new PdfStream(dictionary, [0]), objects, maxImageReferenceSteps: 2));
    }
}
