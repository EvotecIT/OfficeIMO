using OfficeIMO.Pdf;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Fact]
    public void RemovalDeletesWholePairsFromEveryNameTreeNamesArray() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] withCustomTree = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray associations = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, associations, "content-credential.c2pa");
            var pairs = new PdfArray();
            pairs.Items.Add(new PdfStringObj("credential", true));
            pairs.Items.Add(candidate);
            var customTree = new PdfDictionary();
            customTree.Items["Names"] = pairs;
            catalog.Items["CustomNameTree"] = customTree;
            return security.InfoObjectNumber;
        });

        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(withCustomTree);
        var parsed = PdfSyntax.ParseObjects(result.ToArray());
        PdfDictionary catalog = Assert.IsType<PdfDictionary>(PdfSyntax.FindCatalog(parsed.Map, parsed.TrailerRaw));
        PdfDictionary tree = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(parsed.Map, catalog.Items["CustomNameTree"]));

        Assert.False(tree.Items.ContainsKey("Names"));
    }

    [Fact]
    public void RemovalDeletesRepliesToRemovedFileAttachmentAnnotations() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        int attachmentNumber = 0;
        int replyNumber = 0;
        byte[] withReply = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray associations = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, associations, "content-credential.c2pa");
            PdfDictionary page = Assert.IsType<PdfDictionary>(objects.Values.Select(item => item.Value)
                .First(value => value is PdfDictionary dictionary && dictionary.Get<PdfName>("Type")?.Name == "Page"));
            attachmentNumber = objects.Keys.Max() + 1;
            replyNumber = attachmentNumber + 1;
            var attachment = new PdfDictionary();
            attachment.Items["Type"] = new PdfName("Annot");
            attachment.Items["Subtype"] = new PdfName("FileAttachment");
            attachment.Items["FS"] = candidate;
            var reply = new PdfDictionary();
            reply.Items["Type"] = new PdfName("Annot");
            reply.Items["Subtype"] = new PdfName("Text");
            reply.Items["IRT"] = new PdfReference(attachmentNumber, 0);
            objects[attachmentNumber] = new PdfIndirectObject(attachmentNumber, 0, attachment);
            objects[replyNumber] = new PdfIndirectObject(replyNumber, 0, reply);
            var annotations = new PdfArray();
            annotations.Items.Add(new PdfReference(attachmentNumber, 0));
            annotations.Items.Add(new PdfReference(replyNumber, 0));
            page.Items["Annots"] = annotations;
            return security.InfoObjectNumber;
        });

        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(withReply);
        var parsed = PdfSyntax.ParseObjects(result.ToArray());

        Assert.DoesNotContain(attachmentNumber, parsed.Map.Keys);
        Assert.DoesNotContain(replyNumber, parsed.Map.Keys);
    }
}
