using OfficeIMO.Pdf;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Theory]
    [InlineData("BS")]
    [InlineData("BE")]
    public void ActiveAnnotationStyleDictionariesCannotOwnProvenanceAssociations(string styleKey) {
        AssertStructuralOwnerRejected((objects, _, candidate) => {
            PdfDictionary candidateDictionary = Assert.IsType<PdfDictionary>(objects[candidate.ObjectNumber].Value);
            candidateDictionary.Items.Remove("Type");
            PdfDictionary page = Assert.Single(
                objects.Values.Select(item => item.Value).OfType<PdfDictionary>(),
                dictionary => dictionary.Get<PdfName>("Type")?.Name == "Page");
            var annotation = new PdfDictionary();
            annotation.Items["Type"] = new PdfName("Annot");
            annotation.Items["Subtype"] = new PdfName("Text");
            annotation.Items[styleKey] = candidate;
            page.Items["Annots"] = ArrayWith(AddObject(objects, annotation));
        });
    }

    [Fact]
    public void ActiveCatalogUriDictionaryCannotOwnProvenanceAssociations() {
        AssertStructuralOwnerRejected((objects, catalog, candidate) => {
            PdfDictionary candidateDictionary = Assert.IsType<PdfDictionary>(objects[candidate.ObjectNumber].Value);
            candidateDictionary.Items.Remove("Type");
            catalog.Items["URI"] = candidate;
        });
    }

    [Fact]
    public void SharedAnnotationAppearanceGraphsRemainStructural() {
        AssertStructuralOwnerRejected((objects, _, candidate) => {
            PdfDictionary page = Assert.Single(
                objects.Values.Select(item => item.Value).OfType<PdfDictionary>(),
                dictionary => dictionary.Get<PdfName>("Type")?.Name == "Page");
            var appearance = new PdfDictionary();
            appearance.Items["AF"] = ArrayWith(candidate);
            PdfReference sharedAppearance = AddObject(objects, appearance);
            var annotations = new PdfArray();
            for (int index = 0; index < 256; index++) {
                var annotation = new PdfDictionary();
                annotation.Items["Type"] = new PdfName("Annot");
                annotation.Items["Subtype"] = new PdfName("Text");
                annotation.Items["AP"] = sharedAppearance;
                annotations.Items.Add(AddObject(objects, annotation));
            }
            page.Items["Annots"] = annotations;
        });
    }

    [Fact]
    public void RemovalProcessesReverseOrderedReplyChains() {
        byte[] original = CreatePdfWithCandidateAndRetainedAttachment();
        var removedObjectNumbers = new List<int>();
        byte[] withReplies = PdfDocumentObjectGraphRewriter.Rewrite(original, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray associations = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, associations, "content-credential.c2pa");
            PdfDictionary page = Assert.Single(
                objects.Values.Select(item => item.Value).OfType<PdfDictionary>(),
                dictionary => dictionary.Get<PdfName>("Type")?.Name == "Page");
            int nextObjectNumber = objects.Keys.Max() + 1;
            var attachment = new PdfDictionary();
            attachment.Items["Type"] = new PdfName("Annot");
            attachment.Items["Subtype"] = new PdfName("FileAttachment");
            attachment.Items["FS"] = candidate;
            int parentNumber = nextObjectNumber++;
            objects[parentNumber] = new PdfIndirectObject(parentNumber, 0, attachment);
            removedObjectNumbers.Add(parentNumber);
            for (int index = 0; index < 512; index++) {
                var reply = new PdfDictionary();
                reply.Items["Type"] = new PdfName("Annot");
                reply.Items["Subtype"] = new PdfName("Text");
                reply.Items["IRT"] = new PdfReference(parentNumber, 0);
                parentNumber = nextObjectNumber++;
                objects[parentNumber] = new PdfIndirectObject(parentNumber, 0, reply);
                removedObjectNumbers.Add(parentNumber);
            }
            var annotations = new PdfArray();
            foreach (int objectNumber in removedObjectNumbers.AsEnumerable().Reverse()) {
                annotations.Items.Add(new PdfReference(objectNumber, 0));
            }
            page.Items["Annots"] = annotations;
            return security.InfoObjectNumber;
        });

        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(withReplies);
        var parsed = PdfSyntax.ParseObjects(result.ToArray());

        Assert.All(removedObjectNumbers, objectNumber => Assert.DoesNotContain(objectNumber, parsed.Map.Keys));
    }
}
