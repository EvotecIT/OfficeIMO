using System.Text;
using OfficeIMO.Pdf;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Fact]
    public void ProvenanceInspectionRejectsUndecodableCandidateAttachmentFilters() {
        byte[] pdf = RewriteCandidateAssociation((objects, catalog, candidate) => {
            PdfDictionary fileSpecification = Assert.IsType<PdfDictionary>(objects[candidate.ObjectNumber].Value);
            PdfDictionary embeddedFiles = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, fileSpecification.Items["EF"]));
            PdfStream stream = Assert.IsType<PdfStream>(PdfObjectLookup.Resolve(objects, embeddedFiles.Items["F"]));
            stream.Dictionary.Items["Filter"] = new PdfName("DCTDecode");
            catalog.Items["AF"] = ArrayWith(candidate);
        });

        Assert.Throws<InvalidDataException>(() => PdfProvenance.Inspect(pdf));
    }
}
