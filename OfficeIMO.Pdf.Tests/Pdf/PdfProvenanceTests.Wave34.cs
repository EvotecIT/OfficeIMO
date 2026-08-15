using System.Text;
using OfficeIMO.Pdf;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Fact]
    public void CatalogXmpConsumesTheSharedExpandedByteBudgetDuringOpen() {
        byte[] metadata = Enumerable.Repeat((byte)'x', 256).ToArray();
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.7\n" +
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /Metadata 4 0 R >>\nendobj\n" +
            "2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n" +
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 10 10] >>\nendobj\n" +
            "4 0 obj\n<< /Type /Metadata /Subtype /XML /Length " + metadata.Length + " >>\nstream\n");
        output.Write(metadata, 0, metadata.Length);
        WriteAscii(output, "\nendstream\nendobj\ntrailer\n<< /Root 1 0 R >>\n%%EOF\n");

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => PdfProvenance.Inspect(
            output.ToArray(),
            new OfficeProvenanceOptions { MaxExpandedContainerBytes = 128 }));

        Assert.Equal(PdfReadLimitKind.DecodedStreamBytes, exception.Kind);
    }

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
