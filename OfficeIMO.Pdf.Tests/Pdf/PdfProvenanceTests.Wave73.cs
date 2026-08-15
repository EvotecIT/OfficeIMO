using OfficeIMO.Pdf;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Fact]
    public void ExplicitLowerAggregateDecodeLimitRemainsAuthoritative() {
        const long explicitLimit = 32L * 1024L * 1024L;
        var options = new OfficeProvenanceOptions { MaxExpandedContainerBytes = 512L * 1024L * 1024L };
        var requested = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxTotalDecodedStreamBytes = explicitLimit }
        };

        PdfReadOptions effective = PdfProvenance.CreateReadOptionsForInspection(options, requested);

        Assert.Equal(explicitLimit, effective.Limits.MaxTotalDecodedStreamBytes);
    }

    [Fact]
    public void ExplicitHigherAggregateDecodeLimitIsCappedByProvenanceLimit() {
        const long provenanceLimit = 256L * 1024L * 1024L;
        var options = new OfficeProvenanceOptions { MaxExpandedContainerBytes = provenanceLimit };
        var requested = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxTotalDecodedStreamBytes = 768L * 1024L * 1024L }
        };

        PdfReadOptions effective = PdfProvenance.CreateReadOptionsForInspection(options, requested);

        Assert.Equal(provenanceLimit, effective.Limits.MaxTotalDecodedStreamBytes);
    }

    [Fact]
    public void ActiveThreeDimensionalAnimationGraphsCannotOwnProvenanceAssociations() {
        AssertStructuralOwnerRejected((objects, catalog, candidate) => {
            var streamDictionary = new PdfDictionary();
            streamDictionary.Items["Type"] = new PdfName("3D");
            streamDictionary.Items["AN"] = candidate;
            catalog.Items["Private3DStream"] = AddObject(objects,
                new PdfStream(streamDictionary, Array.Empty<byte>()));
        });
    }

    [Theory]
    [InlineData("Measure")]
    [InlineData("ExData")]
    public void ActiveAnnotationStructuralGraphsCannotOwnProvenanceAssociations(string key) {
        AssertStructuralOwnerRejected((objects, _, candidate) => {
            PdfDictionary page = Assert.Single(objects.Values.Select(item => item.Value).OfType<PdfDictionary>(),
                dictionary => dictionary.Get<PdfName>("Type")?.Name == "Page");
            var annotation = new PdfDictionary();
            annotation.Items["Type"] = new PdfName("Annot");
            annotation.Items["Subtype"] = new PdfName("Text");
            annotation.Items[key] = candidate;
            page.Items["Annots"] = ArrayWith(AddObject(objects, annotation));
        });
    }
}
