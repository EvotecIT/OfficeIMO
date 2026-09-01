using OfficeIMO.Pdf;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Fact]
    public void ExplicitHigherPerStreamDecodeLimitIsNotClampedToManifestLimit() {
        const int manifestLimit = 64 * 1024 * 1024;
        const int explicitLimit = 128 * 1024 * 1024;
        var options = new OfficeProvenanceOptions {
            MaxAssetBytes = 256L * 1024L * 1024L,
            MaxManifestBytes = manifestLimit,
            MaxExpandedContainerBytes = 256L * 1024L * 1024L
        };
        var requested = new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxDecodedStreamBytes = explicitLimit }
        };

        PdfLoadOptions effective = PdfProvenance.CreateReadOptionsForInspection(options, requested);

        Assert.Equal(explicitLimit, effective.Limits.MaxDecodedStreamBytes);
    }
}
