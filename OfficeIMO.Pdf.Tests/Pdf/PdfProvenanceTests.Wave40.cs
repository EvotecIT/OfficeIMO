using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Fact]
    public void ActiveWidgetAppearanceCharacteristicsCannotOwnProvenanceAssociations() {
        AssertStructuralOwnerRejected((objects, _, candidate) => {
            PdfDictionary page = Assert.Single(
                objects.Values.Select(item => item.Value).OfType<PdfDictionary>(),
                dictionary => dictionary.Get<PdfName>("Type")?.Name == "Page");
            var widget = new PdfDictionary();
            widget.Items["Type"] = new PdfName("Annot");
            widget.Items["Subtype"] = new PdfName("Widget");
            widget.Items["MK"] = candidate;
            page.Items["Annots"] = ArrayWith(AddObject(objects, widget));
        });
    }

    [Fact]
    public void ActiveResourceStreamDecodeParametersCannotOwnProvenanceAssociations() {
        AssertStructuralOwnerRejected((objects, _, candidate) => {
            PdfDictionary page = Assert.Single(
                objects.Values.Select(item => item.Value).OfType<PdfDictionary>(),
                dictionary => dictionary.Get<PdfName>("Type")?.Name == "Page");
            var imageDictionary = new PdfDictionary();
            imageDictionary.Items["Type"] = new PdfName("XObject");
            imageDictionary.Items["Subtype"] = new PdfName("Image");
            imageDictionary.Items["Width"] = new PdfNumber(1);
            imageDictionary.Items["Height"] = new PdfNumber(1);
            imageDictionary.Items["ColorSpace"] = new PdfName("DeviceGray");
            imageDictionary.Items["BitsPerComponent"] = new PdfNumber(8);
            imageDictionary.Items["DecodeParms"] = candidate;
            PdfReference image = AddObject(objects, new PdfStream(imageDictionary, new byte[] { 0 }));
            var xObjects = new PdfDictionary();
            xObjects.Items["Im1"] = image;
            var resources = new PdfDictionary();
            resources.Items["XObject"] = xObjects;
            page.Items["Resources"] = resources;
        });
    }

    [Fact]
    public void ActiveCatalogPermissionsDictionaryCannotOwnProvenanceAssociations() {
        AssertStructuralOwnerRejected((_, catalog, candidate) => catalog.Items["Perms"] = candidate);
    }
}
