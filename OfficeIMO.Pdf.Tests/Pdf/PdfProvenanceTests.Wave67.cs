using OfficeIMO.Pdf;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Fact]
    public void SynthesizedReadOptionsHonorRaisedPerManifestDecodeLimit() {
        const long manifestLimit = 300L * 1024L * 1024L;
        var options = new OfficeProvenanceOptions {
            MaxAssetBytes = 600L * 1024L * 1024L,
            MaxManifestBytes = manifestLimit,
            MaxExpandedContainerBytes = 600L * 1024L * 1024L
        };

        PdfReadOptions effective = PdfProvenance.CreateReadOptionsForInspection(options, readOptions: null);

        Assert.Equal(manifestLimit, effective.Limits.MaxDecodedStreamBytes);
    }

    [Fact]
    public void ExplicitLowerPerStreamDecodeLimitRemainsAuthoritative() {
        const int explicitLimit = 128 * 1024 * 1024;
        var options = new OfficeProvenanceOptions {
            MaxAssetBytes = 600L * 1024L * 1024L,
            MaxManifestBytes = 300L * 1024L * 1024L,
            MaxExpandedContainerBytes = 600L * 1024L * 1024L
        };
        var requested = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxDecodedStreamBytes = explicitLimit }
        };

        PdfReadOptions effective = PdfProvenance.CreateReadOptionsForInspection(options, requested);

        Assert.Equal(explicitLimit, effective.Limits.MaxDecodedStreamBytes);
    }
}
