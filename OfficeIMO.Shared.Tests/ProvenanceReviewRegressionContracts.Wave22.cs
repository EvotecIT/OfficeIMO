using System.Text;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceReviewRegressionContracts {
    [Fact]
    public void GifXmpSubBlocksRespectTheSharedContainerEntryBudget() {
        byte[] packet = Encoding.UTF8.GetBytes(new string('x', 2048));
        byte[] gif = CreateGifWithSubBlockedXmp(packet);
        var options = new OfficeProvenanceOptions { MaxContainerEntries = 2 };

        Assert.Throws<InvalidDataException>(() => OfficeProvenanceInspector.Inspect(gif, "fixture.gif", options));
    }
}
