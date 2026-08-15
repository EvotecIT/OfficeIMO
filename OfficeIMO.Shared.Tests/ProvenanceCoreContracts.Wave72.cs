using System.Text;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Fact]
    public void OversizedSvgManifestCannotBeRemovedPermissively() {
        string encoded = Convert.ToBase64String(CreateManifestStore());
        byte[] svg = Encoding.UTF8.GetBytes(
            "<svg xmlns='http://www.w3.org/2000/svg' xmlns:c2pa='http://c2pa.org/manifest'>" +
            "<metadata><c2pa:manifest>" + encoded + "</c2pa:manifest></metadata></svg>");
        var options = new OfficeProvenanceRemovalOptions { RequireStructurallyValidCarrier = false };
        options.Limits.MaxManifestBytes = 1;

        Assert.Throws<InvalidDataException>(() =>
            OfficeProvenanceRemover.Remove(svg, "fixture.svg", options));
    }
}
