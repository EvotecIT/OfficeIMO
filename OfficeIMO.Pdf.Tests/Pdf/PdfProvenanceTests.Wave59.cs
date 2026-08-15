using OfficeIMO.Pdf;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Fact]
    public void CatalogPieceInfoGraphsCannotOwnProvenanceAssociations() {
        AssertStructuralOwnerRejected((_, catalog, candidate) => catalog.Items["PieceInfo"] = candidate);
    }

    [Fact]
    public void ActivePageBoxColorInfoGraphsCannotOwnProvenanceAssociations() {
        AssertStructuralOwnerRejected((objects, _, candidate) => {
            PdfDictionary page = Assert.Single(
                objects.Values.Select(item => item.Value).OfType<PdfDictionary>(),
                dictionary => dictionary.Get<PdfName>("Type")?.Name == "Page");
            page.Items["BoxColorInfo"] = candidate;
        });
    }

    [Fact]
    public void ProvenanceAggregateAttachmentLimitCanExceedTheGenericDefault() {
        const long provenanceLimit = 300L * 1024L * 1024L;

        PdfReadOptions adjusted = PdfReadOptions.WithMaximumContainerEntries(
            options: null,
            maximumContainerEntries: 64,
            maximumDecodedStreamBytes: 16L * 1024L * 1024L,
            maximumTotalAttachmentBytes: provenanceLimit);

        Assert.Equal(provenanceLimit, adjusted.Limits.MaxTotalAttachmentBytes);
    }

    [Fact]
    public void FullRewritePreservesPermanentTrailerIdentifierAndRegeneratesRevisionIdentifier() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        (byte[] permanentBefore, string identifiersBefore) = ReadTrailerIdentifiers(pdf);

        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(pdf);
        (byte[] permanentAfter, string identifiersAfter) = ReadTrailerIdentifiers(result.ToArray());

        Assert.True(result.WasChanged);
        Assert.Equal(permanentBefore, permanentAfter);
        Assert.NotEqual(identifiersBefore, identifiersAfter);
    }

    private static (byte[] Permanent, string Identifiers) ReadTrailerIdentifiers(byte[] pdf) {
        string trailer = PdfSyntax.ParseObjects(pdf).TrailerRaw;
        return (
            Assert.IsType<byte[]>(PdfSyntax.ReadPermanentTrailerIdentifier(trailer)),
            PdfIncrementalObjectWriter.ReadTrailerIdEntry(trailer));
    }
}
