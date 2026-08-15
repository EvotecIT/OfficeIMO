using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Fact]
    public void ActiveCatalogSpiderInfoGraphsCannotOwnProvenanceAssociations() {
        AssertStructuralOwnerRejected((_, catalog, candidate) => catalog.Items["SpiderInfo"] = candidate);
    }

    [Fact]
    public void ActivePagePresentationStepGraphsCannotOwnProvenanceAssociations() {
        AssertStructuralOwnerRejected((objects, _, candidate) => {
            PdfDictionary page = Assert.Single(
                objects.Values.Select(item => item.Value).OfType<PdfDictionary>(),
                dictionary => dictionary.Get<PdfName>("Type")?.Name == "Page");
            page.Items["PresSteps"] = candidate;
        });
    }

    [Fact]
    public void ActiveImageAlternateGraphsCannotOwnProvenanceAssociations() {
        AssertStructuralOwnerRejected((objects, catalog, candidate) => {
            var image = new PdfDictionary();
            image.Items["Type"] = new PdfName("XObject");
            image.Items["Subtype"] = new PdfName("Image");
            var alternates = new PdfArray();
            alternates.Items.Add(candidate);
            image.Items["Alternates"] = alternates;
            catalog.Items["PrivateImage"] = AddObject(objects, new PdfStream(image, Array.Empty<byte>()));
        });
    }

    [Fact]
    public void ActiveMovieAnnotationGraphsCannotOwnProvenanceAssociations() {
        AssertStructuralOwnerRejected((objects, _, candidate) => {
            PdfDictionary page = Assert.Single(
                objects.Values.Select(item => item.Value).OfType<PdfDictionary>(),
                dictionary => dictionary.Get<PdfName>("Type")?.Name == "Page");
            var annotation = new PdfDictionary();
            annotation.Items["Type"] = new PdfName("Annot");
            annotation.Items["Subtype"] = new PdfName("Movie");
            annotation.Items["Movie"] = candidate;
            var annotations = new PdfArray();
            annotations.Items.Add(AddObject(objects, annotation));
            page.Items["Annots"] = annotations;
        });
    }
}
