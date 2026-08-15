using OfficeIMO.Pdf;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Theory]
    [InlineData("OpenAction")]
    [InlineData("AA")]
    public void ActiveCatalogActionsCannotMasqueradeAsFileSpecifications(string catalogKey) {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] structuralCarrier = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray associations = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, associations, "content-credential.c2pa");
            PdfDictionary action = Assert.IsType<PdfDictionary>(objects[candidate.ObjectNumber].Value);
            action.Items.Remove("Type");
            action.Items["S"] = new PdfName("JavaScript");
            if (catalogKey == "OpenAction") {
                catalog.Items[catalogKey] = candidate;
            } else {
                var additionalActions = new PdfDictionary();
                additionalActions.Items["WC"] = candidate;
                catalog.Items[catalogKey] = additionalActions;
            }
            return security.InfoObjectNumber;
        });

        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(structuralCarrier);

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
    }
}
