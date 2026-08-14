using OfficeIMO.Pdf;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Fact]
    public void RemovalResolvesIndirectFileAttachmentSubtypes() {
        byte[] original = CreatePdfWithCandidateAndRetainedAttachment();
        int annotationNumber = 0;
        byte[] pdf = PdfDocumentObjectGraphRewriter.Rewrite(original, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray associations = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, associations, "content-credential.c2pa");
            PdfDictionary page = Assert.Single(objects.Values.Select(item => item.Value).OfType<PdfDictionary>(),
                dictionary => dictionary.Get<PdfName>("Type")?.Name == "Page");
            var annotation = new PdfDictionary();
            annotation.Items["Type"] = new PdfName("Annot");
            annotation.Items["Subtype"] = AddObject(objects, new PdfName("FileAttachment"));
            annotation.Items["FS"] = candidate;
            PdfReference reference = AddObject(objects, annotation);
            annotationNumber = reference.ObjectNumber;
            page.Items["Annots"] = ArrayWith(reference);
            return security.InfoObjectNumber;
        });

        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(pdf);
        var parsed = PdfSyntax.ParseObjects(result.ToArray());

        Assert.DoesNotContain(annotationNumber, parsed.Map.Keys);
    }

    [Fact]
    public void ActiveFormReferenceFileSpecificationGraphIsProtected() {
        AssertStructuralOwnerRejected((objects, _, candidate) => {
            PdfDictionary page = Assert.Single(objects.Values.Select(item => item.Value).OfType<PdfDictionary>(),
                dictionary => dictionary.Get<PdfName>("Type")?.Name == "Page");
            var reference = new PdfDictionary();
            reference.Items["F"] = candidate;
            var formDictionary = new PdfDictionary();
            formDictionary.Items["Type"] = new PdfName("XObject");
            formDictionary.Items["Subtype"] = new PdfName("Form");
            formDictionary.Items["Ref"] = AddObject(objects, reference);
            var xObjects = new PdfDictionary();
            xObjects.Items["Fm1"] = AddObject(objects, new PdfStream(formDictionary, Array.Empty<byte>()));
            var resources = new PdfDictionary();
            resources.Items["XObject"] = xObjects;
            page.Items["Resources"] = resources;
        });
    }
}
