using OfficeIMO.Pdf;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Fact]
    public void ActivePagePieceInfoCannotMasqueradeAsFileSpecification() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] structuralCarrier = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray associations = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, associations, "content-credential.c2pa");
            PdfDictionary fileSpecification = Assert.IsType<PdfDictionary>(objects[candidate.ObjectNumber].Value);
            fileSpecification.Items.Remove("Type");
            PdfDictionary page = Assert.IsType<PdfDictionary>(objects.Values
                .Select(item => item.Value)
                .First(value => value is PdfDictionary dictionary && dictionary.Get<PdfName>("Type")?.Name == "Page"));
            page.Items["PieceInfo"] = candidate;
            return security.InfoObjectNumber;
        });

        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(structuralCarrier);

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void LegacyDestinationsStopAtReferencedPages() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] withDestination = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfReference pageReference = objects.Values
                .Where(item => item.Value is PdfDictionary dictionary && dictionary.Get<PdfName>("Type")?.Name == "Page")
                .Select(item => new PdfReference(item.ObjectNumber, item.Generation))
                .First();
            var destination = new PdfArray();
            destination.Items.Add(pageReference);
            destination.Items.Add(new PdfName("Fit"));
            var destinations = new PdfDictionary();
            destinations.Items["Start"] = destination;
            catalog.Items["Dests"] = destinations;
            return security.InfoObjectNumber;
        });

        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(withDestination);

        Assert.True(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void FileSpecificationRequiresAnActualFilename() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] missingName = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray associations = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, associations, "content-credential.c2pa");
            PdfDictionary fileSpecification = Assert.IsType<PdfDictionary>(objects[candidate.ObjectNumber].Value);
            fileSpecification.Items.Remove("F");
            fileSpecification.Items.Remove("UF");
            return security.InfoObjectNumber;
        });

        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(missingName);

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
    }
}
