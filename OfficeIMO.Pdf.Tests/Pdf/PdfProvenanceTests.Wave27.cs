using OfficeIMO.Pdf;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Fact]
    public void ActivePortfolioCollectionCannotMasqueradeAsFileSpecification() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] structuralCarrier = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray associations = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, associations, "content-credential.c2pa");
            PdfDictionary fileSpecification = Assert.IsType<PdfDictionary>(objects[candidate.ObjectNumber].Value);
            fileSpecification.Items.Remove("Type");
            fileSpecification.Items["View"] = new PdfName("T");
            catalog.Items["Collection"] = candidate;
            return security.InfoObjectNumber;
        });

        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(structuralCarrier);

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
    }
}
