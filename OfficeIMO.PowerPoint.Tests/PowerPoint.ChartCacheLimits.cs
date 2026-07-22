using System.Reflection;
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
        MethodInfo method = typeof(PowerPointChart)
            .GetMethod("GetCachedPointLength", BindingFlags.NonPublic | BindingFlags.Static)!
            .MakeGenericMethod(typeof(C.NumericPoint));

        TargetInvocationException invocation = Assert.Throws<TargetInvocationException>(() =>
            method.Invoke(null, new object[] { cache, points, getIndex }));

        Assert.IsType<InvalidDataException>(invocation.InnerException);
    }

    [Fact]
    public void ChartSnapshotRejectsSparseOversizedCacheIndexBeforeAllocation() {
        var cache = new C.NumberingCache();
        var points = new List<C.NumericPoint> {
            new C.NumericPoint { Index = 100_000U }
        };
        Func<C.NumericPoint, uint?> getIndex = point => point.Index?.Value;
        MethodInfo method = typeof(PowerPointChart)
            .GetMethod("GetCachedPointLength", BindingFlags.NonPublic | BindingFlags.Static)!
            .MakeGenericMethod(typeof(C.NumericPoint));

        TargetInvocationException invocation = Assert.Throws<TargetInvocationException>(() =>
            method.Invoke(null, new object[] { cache, points, getIndex }));

        Assert.IsType<InvalidDataException>(invocation.InnerException);
    }
}
