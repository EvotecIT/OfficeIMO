using System.Text;
using OfficeIMO.Pdf;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Fact]
    public void CandidateStreamsMustDeclareTheEmbeddedFileType() {
        byte[] pdf = RewriteCandidateAssociation((objects, catalog, candidate) => {
            PdfDictionary fileSpecification = Assert.IsType<PdfDictionary>(objects[candidate.ObjectNumber].Value);
            PdfDictionary embeddedFiles = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, fileSpecification.Items["EF"]));
            PdfStream stream = Assert.IsType<PdfStream>(PdfObjectLookup.Resolve(objects, embeddedFiles.Items["F"]));
            stream.Dictionary.Items.Remove("Type");
            catalog.Items["AF"] = ArrayWith(candidate);
        });

        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(pdf);

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void ActivePageTransitionDictionariesCannotOwnProvenanceAssociations() {
        AssertStructuralOwnerRejected((objects, _, candidate) => {
            PdfDictionary page = Assert.Single(
                objects.Values.Select(item => item.Value).OfType<PdfDictionary>(),
                dictionary => dictionary.Get<PdfName>("Type")?.Name == "Page");
            var transition = new PdfDictionary();
            transition.Items["S"] = new PdfName("Dissolve");
            transition.Items["AF"] = ArrayWith(candidate);
            page.Items["Trans"] = AddObject(objects, transition);
        });
    }

    [Fact]
    public void OutputIntentProfilesHonorTheSharedExpandedByteBudgetDuringOpen() {
        byte[] profile = Enumerable.Repeat((byte)'x', 256).ToArray();
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.7\n" +
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /OutputIntents [4 0 R] >>\nendobj\n" +
            "2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n" +
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 10 10] >>\nendobj\n" +
            "4 0 obj\n<< /Type /OutputIntent /S /GTS_PDFA1 /DestOutputProfile 5 0 R >>\nendobj\n" +
            "5 0 obj\n<< /N 3 /Length " + profile.Length + " >>\nstream\n");
        output.Write(profile, 0, profile.Length);
        WriteAscii(output, "\nendstream\nendobj\ntrailer\n<< /Root 1 0 R >>\n%%EOF\n");

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => PdfProvenance.Inspect(
            output.ToArray(),
            new OfficeProvenanceOptions { MaxExpandedContainerBytes = 128 }));

        Assert.Equal(PdfReadLimitKind.DecodedStreamBytes, exception.Kind);
    }
}
