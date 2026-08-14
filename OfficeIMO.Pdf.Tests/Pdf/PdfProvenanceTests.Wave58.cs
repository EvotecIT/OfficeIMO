using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Theory]
    [InlineData("Form")]
    [InlineData("Image")]
    public void ActiveXObjectOptionalContentGraphIsProtected(string subtype) {
        AssertStructuralOwnerRejected((objects, _, candidate) => {
            PdfDictionary page = Assert.Single(objects.Values.Select(item => item.Value).OfType<PdfDictionary>(),
                dictionary => dictionary.Get<PdfName>("Type")?.Name == "Page");
            var optionalContent = new PdfDictionary();
            optionalContent.Items["F"] = candidate;
            var xObjectDictionary = new PdfDictionary();
            xObjectDictionary.Items["Type"] = new PdfName("XObject");
            xObjectDictionary.Items["Subtype"] = new PdfName(subtype);
            xObjectDictionary.Items["OC"] = AddObject(objects, optionalContent);
            var xObjects = new PdfDictionary();
            xObjects.Items["X1"] = AddObject(objects, new PdfStream(xObjectDictionary, Array.Empty<byte>()));
            var resources = new PdfDictionary();
            resources.Items["XObject"] = xObjects;
            page.Items["Resources"] = resources;
        });
    }
}
