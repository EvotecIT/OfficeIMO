using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using Xunit;

namespace OfficeIMO.Tests;

public partial class PowerPoint {
    [Fact]
    public void MasterTextStyle9_ThreeDuplicateInstancesRemainAmbiguous() {
        LegacyPptRecord[] records = {
            CreateStyle9Record(7, 0),
            CreateStyle9Record(7, 8),
            CreateStyle9Record(7, 16)
        };
        int duplicateCount = 0;

        IReadOnlyDictionary<ushort, LegacyPptRecord> selected =
            LegacyPptPresentation.CollectUniqueMasterTextStyle9Records(
                records,
                _ => duplicateCount++);

        Assert.Empty(selected);
        Assert.Equal(1, duplicateCount);
    }

    private static LegacyPptRecord CreateStyle9Record(ushort instance, int offset) {
        byte[] source = new byte[24];
        return new LegacyPptRecord(
            source,
            offset,
            version: 0,
            instance,
            type: 0x0FAD,
            payloadOffset: offset + 8,
            payloadLength: 0);
    }
}
