using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Fact]
    public void ActiveImageColorSpaceGraphsCannotOwnProvenanceAssociations() {
        AssertStructuralOwnerRejected((objects, _, candidate) => {
            PdfDictionary page = Assert.Single(
                objects.Values.Select(item => item.Value).OfType<PdfDictionary>(),
                dictionary => dictionary.Get<PdfName>("Type")?.Name == "Page");
            var colorSpace = new PdfArray();
            colorSpace.Items.Add(new PdfName("Separation"));
            colorSpace.Items.Add(new PdfName("Spot"));
            colorSpace.Items.Add(new PdfName("DeviceRGB"));
            colorSpace.Items.Add(candidate);
            var imageDictionary = new PdfDictionary();
            imageDictionary.Items["Type"] = new PdfName("XObject");
            imageDictionary.Items["Subtype"] = new PdfName("Image");
            imageDictionary.Items["Width"] = new PdfNumber(1);
            imageDictionary.Items["Height"] = new PdfNumber(1);
            imageDictionary.Items["ColorSpace"] = colorSpace;
            imageDictionary.Items["BitsPerComponent"] = new PdfNumber(8);
            PdfReference image = AddObject(objects, new PdfStream(imageDictionary, new byte[] { 0 }));
            var xObjects = new PdfDictionary();
            xObjects.Items["Im1"] = image;
            var resources = new PdfDictionary();
            resources.Items["XObject"] = xObjects;
            page.Items["Resources"] = resources;
        });
    }

    [Fact]
    public void ProvenanceRemovalRefusesToSilentlyDropPdfEncryption() {
        byte[] encrypted = PdfDocument.Create(new PdfOptions().SetEncryption("open", "owner"))
            .Paragraph(paragraph => paragraph.Text("Encrypted provenance"))
            .AttachFile(new PdfEmbeddedFile(
                "content-credential.c2pa",
                CreateManifestStore(),
                "application/c2pa",
                PdfAssociatedFileRelationship.C2paManifest))
            .ToBytes();
        var readOptions = new PdfLoadOptions { Password = "owner" };

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            PdfProvenance.Remove(encrypted, readOptions: readOptions));

        Assert.Contains("encryption", exception.Message, StringComparison.OrdinalIgnoreCase);
    }
}
