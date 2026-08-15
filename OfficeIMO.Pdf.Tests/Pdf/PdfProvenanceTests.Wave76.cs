using OfficeIMO.Pdf;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Fact]
    public void EmptyFilterArrayRequiredDecodeReusesThePermissiveBudgetEntry() {
        var budget = new PdfDecodedStreamBudget(new PdfReadLimits {
            MaxDecodedStreamBytes = 8,
            MaxTotalDecodedStreamBytes = 5
        });
        var dictionary = new PdfDictionary();
        dictionary.Items["Filter"] = new PdfArray();
        var stream = new PdfStream(dictionary, new byte[] { 1, 2, 3 });
        var objects = new Dictionary<int, PdfIndirectObject>();

        byte[] permissive = budget.Decode(stream, objects, maximumRequestedBytes: 8);
        byte[] required = budget.DecodeRequired(stream, objects, maximumRequestedBytes: 8);

        Assert.Same(permissive, required);
        Assert.Equal(3, budget.UsedBytes);
    }

    [Fact]
    public void SynthesizedReadOptionsHonorRaisedRawManifestLimit() {
        const long manifestLimit = 300L * 1024L * 1024L;
        var options = new OfficeProvenanceOptions {
            MaxAssetBytes = 600L * 1024L * 1024L,
            MaxManifestBytes = manifestLimit,
            MaxExpandedContainerBytes = 600L * 1024L * 1024L
        };

        PdfReadOptions effective = PdfProvenance.CreateReadOptionsForInspection(options, readOptions: null);

        Assert.Equal(manifestLimit, effective.Limits.MaxRawStreamBytes);
    }

    [Fact]
    public void ExplicitLowerRawStreamLimitRemainsAuthoritative() {
        const int explicitLimit = 128 * 1024 * 1024;
        var options = new OfficeProvenanceOptions {
            MaxAssetBytes = 600L * 1024L * 1024L,
            MaxManifestBytes = 300L * 1024L * 1024L,
            MaxExpandedContainerBytes = 600L * 1024L * 1024L
        };
        var requested = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxRawStreamBytes = explicitLimit }
        };

        PdfReadOptions effective = PdfProvenance.CreateReadOptionsForInspection(options, requested);

        Assert.Equal(explicitLimit, effective.Limits.MaxRawStreamBytes);
    }
}
