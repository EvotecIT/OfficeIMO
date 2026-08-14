using System.Text;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Fact]
    public void MalformedCompetingGifXmpApplicationInvalidatesTheCarrierSet() {
        byte[] malformed = CreateGifApplication(
            "XMP DataXMP",
            Array.Empty<byte>(),
            Encoding.ASCII.GetBytes("not-an-xmp-application"));
        byte[] valid = CreateGifXmpExtension(CreateXmpPacket());
        byte[] gif = Join(
            Encoding.ASCII.GetBytes("GIF89a"),
            new byte[7],
            malformed,
            valid,
            new byte[] { 0x3B });

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(gif, "fixture.gif");

        Assert.NotEmpty(result.Before.Evidence);
        Assert.All(result.Before.Evidence, evidence => Assert.False(evidence.IsStructurallyValid));
        Assert.False(result.WasChanged);
        Assert.Equal(gif, result.ToArray());
    }
}
