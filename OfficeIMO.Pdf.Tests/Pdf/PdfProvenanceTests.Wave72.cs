using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Fact]
    public void ActiveThreeDimensionalAnnotationActivationGraphsCannotOwnProvenanceAssociations() {
        AssertStructuralOwnerRejected((objects, _, candidate) => {
            PdfDictionary page = Assert.Single(objects.Values.Select(item => item.Value).OfType<PdfDictionary>(),
                dictionary => dictionary.Get<PdfName>("Type")?.Name == "Page");
            var annotation = new PdfDictionary();
            annotation.Items["Type"] = new PdfName("Annot");
            annotation.Items["Subtype"] = new PdfName("3D");
            annotation.Items["3DA"] = candidate;
            page.Items["Annots"] = ArrayWith(AddObject(objects, annotation));
        });
    }

    [Fact]
    public void ActiveThreeDimensionalStreamViewGraphsCannotOwnProvenanceAssociations() {
        AssertStructuralOwnerRejected((objects, _, candidate) => {
            PdfDictionary page = Assert.Single(objects.Values.Select(item => item.Value).OfType<PdfDictionary>(),
                dictionary => dictionary.Get<PdfName>("Type")?.Name == "Page");
            var streamDictionary = new PdfDictionary();
            streamDictionary.Items["Type"] = new PdfName("3D");
            streamDictionary.Items["VA"] = ArrayWith(candidate);
            PdfReference threeDimensionalStream = AddObject(objects, new PdfStream(streamDictionary, Array.Empty<byte>()));
            var annotation = new PdfDictionary();
            annotation.Items["Type"] = new PdfName("Annot");
            annotation.Items["Subtype"] = new PdfName("3D");
            annotation.Items["3DD"] = threeDimensionalStream;
            page.Items["Annots"] = ArrayWith(AddObject(objects, annotation));
        });
    }

    [Fact]
    public void ActiveAnnotationOptionalContentGraphsCannotOwnProvenanceAssociations() {
        AssertStructuralOwnerRejected((objects, _, candidate) => {
            PdfDictionary page = Assert.Single(objects.Values.Select(item => item.Value).OfType<PdfDictionary>(),
                dictionary => dictionary.Get<PdfName>("Type")?.Name == "Page");
            var annotation = new PdfDictionary();
            annotation.Items["Type"] = new PdfName("Annot");
            annotation.Items["Subtype"] = new PdfName("Text");
            annotation.Items["OC"] = candidate;
            page.Items["Annots"] = ArrayWith(AddObject(objects, annotation));
        });
    }
}
