using DocumentFormat.OpenXml;
using OfficeIMO.PowerPoint;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Xunit;

namespace OfficeIMO.Tests;

public partial class PowerPoint {
    [Fact]
    public void ChartSnapshotRejectsOversizedDeclaredCacheBeforeAllocation() {
        var cache = new C.NumberingCache(new C.PointCount { Val = 100_001U });
        var points = new List<C.NumericPoint>();
        Func<C.NumericPoint, uint?> getIndex = point => point.Index?.Value;
        Assert.Throws<InvalidDataException>(() => OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartCacheReader.GetCachedPointLength(
            OfficeIMO.OpenXml.Internal.OfficeXmlNode.FromSdk(cache), points, getIndex, PowerPointUtils.MaximumSharedChartPoints));
    }

    [Fact]
    public void ChartSnapshotRejectsSparseOversizedCacheIndexBeforeAllocation() {
        var cache = new C.NumberingCache();
        var points = new List<C.NumericPoint> {
            new C.NumericPoint { Index = 100_000U }
        };
        Func<C.NumericPoint, uint?> getIndex = point => point.Index?.Value;
        Assert.Throws<InvalidDataException>(() => OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartCacheReader.GetCachedPointLength(
            OfficeIMO.OpenXml.Internal.OfficeXmlNode.FromSdk(cache), points, getIndex, PowerPointUtils.MaximumSharedChartPoints));
    }
}
