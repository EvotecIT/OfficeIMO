using OfficeIMO.Pdf;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Fact]
    public void ActiveCatalogRequirementGraphsCannotOwnProvenanceAssociations() {
        AssertStructuralOwnerRejected((_, catalog, candidate) => catalog.Items["Requirements"] = candidate);
    }

    [Fact]
    public void WatermarkFixedPrintGraphsCannotOwnProvenanceAssociations() {
        AssertStructuralOwnerRejected((objects, _, candidate) => {
            PdfDictionary page = Assert.Single(objects.Values.Select(item => item.Value).OfType<PdfDictionary>(),
                dictionary => dictionary.Get<PdfName>("Type")?.Name == "Page");
            var annotation = new PdfDictionary();
            annotation.Items["Type"] = new PdfName("Annot");
            annotation.Items["Subtype"] = new PdfName("Watermark");
            annotation.Items["FixedPrint"] = candidate;
            page.Items["Annots"] = ArrayWith(AddObject(objects, annotation));
        });
    }

    [Fact]
    public void DirectAcroFormDefaultResourcesCannotOwnProvenanceAssociations() {
        AssertStructuralOwnerRejected((_, catalog, candidate) => {
            var acroForm = new PdfDictionary();
            acroForm.Items["DR"] = candidate;
            catalog.Items["AcroForm"] = acroForm;
        });
    }

    [Fact]
    public void AnnotationOnlyCandidatesResolveIndirectSubtypes() {
        byte[] pdf = RewriteCandidateAssociation((objects, catalog, candidate) => {
            catalog.Items.Remove("AF");
            catalog.Items.Remove("Names");
            PdfDictionary page = Assert.Single(objects.Values.Select(item => item.Value).OfType<PdfDictionary>(),
                dictionary => dictionary.Get<PdfName>("Type")?.Name == "Page");
            var annotation = new PdfDictionary();
            annotation.Items["Type"] = new PdfName("Annot");
            annotation.Items["Subtype"] = AddObject(objects, new PdfName("FileAttachment"));
            annotation.Items["FS"] = candidate;
            page.Items["Annots"] = ArrayWith(AddObject(objects, annotation));
        });

        OfficeProvenanceReport report = PdfProvenance.Inspect(pdf);

        Assert.False(Assert.Single(report.Evidence).IsStructurallyValid);
    }
}
