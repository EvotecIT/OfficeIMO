using OfficeIMO.Pdf;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Fact]
    public void IndirectEmbeddedFileTypeIsAcceptedForTheSelectedCarrier() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] indirectType = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray associations = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, associations, "content-credential.c2pa");
            PdfDictionary fileSpecification = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, candidate));
            PdfDictionary embeddedFiles = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, fileSpecification.Items["EF"]));
            PdfStream embeddedFile = Assert.IsType<PdfStream>(PdfObjectLookup.Resolve(objects, embeddedFiles.Items.Values.First()));
            embeddedFile.Dictionary.Items["Type"] = AddObject(objects, new PdfName("EmbeddedFile"));
            return security.InfoObjectNumber;
        });

        OfficeProvenanceReport report = PdfProvenance.Inspect(indirectType);

        Assert.True(Assert.Single(report.Evidence).IsStructurallyValid);
    }

    [Fact]
    public void IndirectImageSubtypeProtectsItsColorSpaceGraph() {
        AssertStructuralOwnerRejected((objects, _, candidate) => {
            PdfDictionary page = Assert.Single(
                objects.Values.Select(item => item.Value).OfType<PdfDictionary>(),
                dictionary => dictionary.Get<PdfName>("Type")?.Name == "Page");
            var imageDictionary = new PdfDictionary();
            imageDictionary.Items["Type"] = new PdfName("XObject");
            imageDictionary.Items["Subtype"] = AddObject(objects, new PdfName("Image"));
            imageDictionary.Items["Width"] = new PdfNumber(1);
            imageDictionary.Items["Height"] = new PdfNumber(1);
            imageDictionary.Items["ColorSpace"] = candidate;
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
    public void IndirectFormSubtypeProtectsItsTransparencyGroupGraph() {
        AssertStructuralOwnerRejected((objects, _, candidate) => {
            PdfDictionary page = Assert.Single(
                objects.Values.Select(item => item.Value).OfType<PdfDictionary>(),
                dictionary => dictionary.Get<PdfName>("Type")?.Name == "Page");
            var formDictionary = new PdfDictionary();
            formDictionary.Items["Type"] = new PdfName("XObject");
            formDictionary.Items["Subtype"] = AddObject(objects, new PdfName("Form"));
            formDictionary.Items["Group"] = candidate;
            PdfReference form = AddObject(objects, new PdfStream(formDictionary, Array.Empty<byte>()));
            var xObjects = new PdfDictionary();
            xObjects.Items["Fm1"] = form;
            var resources = new PdfDictionary();
            resources.Items["XObject"] = xObjects;
            page.Items["Resources"] = resources;
        });
    }
}
