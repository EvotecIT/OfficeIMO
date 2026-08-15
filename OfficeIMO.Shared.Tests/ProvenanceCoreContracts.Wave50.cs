using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Fact]
    public void ZipRewriteSharesInspectionAndSerializationExpansionBudget() {
        byte[] image = CreatePngWithC2paManifest(CreateManifestStore());
        byte[] package = CreateZip(("media/image.png", image));
        var options = new OfficeProvenanceRemovalOptions();
        options.Limits.MaxExpandedContainerBytes = image.LongLength;

        Assert.Throws<InvalidDataException>(() =>
            OfficeProvenanceRemover.Remove(package, "package.zip", options));
    }
}
