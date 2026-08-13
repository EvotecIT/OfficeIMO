using OfficeIMO.Pdf;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Fact]
    public void UntypedPageTreeAssociationsDoNotValidateDocumentCredentials() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] pageAssociated = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray catalogAssociations = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, catalogAssociations, "content-credential.c2pa");
            catalog.Items.Remove("AF");
            PdfReference pagesReference = Assert.IsType<PdfReference>(catalog.Items["Pages"]);
            PdfDictionary pages = Assert.IsType<PdfDictionary>(objects[pagesReference.ObjectNumber].Value);
            pages.Items.Remove("Type");
            var pageAssociations = new PdfArray();
            pageAssociations.Items.Add(candidate);
            pages.Items["AF"] = pageAssociations;
            return security.InfoObjectNumber;
        });

        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(pageAssociated);

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
    }
}
