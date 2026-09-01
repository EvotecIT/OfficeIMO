using OfficeIMO.Pdf;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Fact]
    public void ActiveImageOpiGraphsCannotOwnProvenanceAssociations() {
        AssertStructuralOwnerRejected((objects, catalog, candidate) => {
            var image = new PdfDictionary();
            image.Items["Subtype"] = new PdfName("Image");
            image.Items["OPI"] = candidate;
            catalog.Items["PrivateImage"] = AddObject(objects, new PdfStream(image, Array.Empty<byte>()));
        });
    }

    [Fact]
    public void OrphanAssociatedFileGraphsAreIgnoredBeforeAttachmentDecoding() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] orphaned = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray associations = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, associations, "content-credential.c2pa");
            var orphan = new PdfDictionary();
            orphan.Items["AF"] = ArrayWith(candidate);
            _ = AddObject(objects, orphan);
            catalog.Items.Remove("AF");
            catalog.Items.Remove("Names");
            return security.InfoObjectNumber;
        });

        OfficeProvenanceReport report = PdfProvenance.Inspect(orphaned);

        Assert.Empty(report.Evidence);
    }

    [Fact]
    public void RemovalPreservesOnlyDanglingUnrelatedNameTreeContent() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] malformed = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray associations = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, associations, "content-credential.c2pa");
            associations.Items.Clear();
            associations.Items.Add(candidate);
            PdfDictionary names = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, catalog.Items["Names"]));
            PdfDictionary embeddedFiles = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, names.Items["EmbeddedFiles"]));
            PdfArray entries = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, embeddedFiles.Items["Names"]));
            entries.Items.Clear();
            entries.Items.Add(new PdfStringObj("content-credential.c2pa", true));
            entries.Items.Add(candidate);
            entries.Items.Add(new PdfStringObj("dangling", true));
            return security.InfoObjectNumber;
        });

        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(malformed);
        var parsed = PdfSyntax.ParseObjects(result.ToArray());
        PdfDictionary catalog = Assert.IsType<PdfDictionary>(PdfSyntax.FindCatalog(parsed.Map, parsed.TrailerRaw));
        PdfDictionary names = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(parsed.Map, catalog.Items["Names"]));
        PdfDictionary embeddedFiles = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(parsed.Map, names.Items["EmbeddedFiles"]));
        PdfArray entries = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(parsed.Map, embeddedFiles.Items["Names"]));

        Assert.Equal("dangling", Assert.IsType<PdfStringObj>(Assert.Single(entries.Items)).Value);
    }

    [Fact]
    public void AggregateDecodedStreamBudgetIsPublicAndValidated() {
        var limits = new PdfReadLimits { MaxTotalDecodedStreamBytes = 768L * 1024L * 1024L };
        Assert.Equal(768L * 1024L * 1024L, limits.MaxTotalDecodedStreamBytes);

        var invalidOptions = new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxTotalDecodedStreamBytes = 0 }
        };
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            PdfReadDocument.Open(CreatePdfWithCandidateAndRetainedAttachment(), invalidOptions));
    }
}
