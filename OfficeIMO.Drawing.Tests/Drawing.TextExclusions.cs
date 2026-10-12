using System;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingTextExclusionTests {
    [Fact]
    public void ExclusionsPreserveBandsAboveAndBelowAndChooseAWideUnobstructedSide() {
        var area = new OfficeTextFlowRegion(0, 0, 100, 100);
        var regions = OfficeTextFlowRegions.Exclude(area, new[] { new OfficeTextFlowRegion(35, 25, 30, 40) }, () => { }, default);
        Assert.Equal(3, regions.Count);
        Assert.Equal((0D, 0D, 100D, 25D), (regions[0].X, regions[0].Y, regions[0].Width, regions[0].Height));
        Assert.Equal((0D, 25D, 35D, 40D), (regions[1].X, regions[1].Y, regions[1].Width, regions[1].Height));
        Assert.Equal((0D, 65D, 100D, 35D), (regions[2].X, regions[2].Y, regions[2].Width, regions[2].Height));
    }

    [Fact]
    public void OverlappingAndOffFrameObjectsAreClippedAndTheirUnionLeavesNoTextInsideTheMask() {
        var area = new OfficeTextFlowRegion(10, 20, 100, 80);
        var regions = OfficeTextFlowRegions.Exclude(area, new[] {
            new OfficeTextFlowRegion(-10, 10, 90, 60), new OfficeTextFlowRegion(60, 10, 80, 60),
            new OfficeTextFlowRegion(300, 300, 20, 20)
        }, () => { }, default);
        var region = Assert.Single(regions);
        Assert.Equal((10D, 70D, 100D, 30D), (region.X, region.Y, region.Width, region.Height));
    }

    [Fact]
    public void RegionConstructionChargesWorkAndStopsOnCancellationOrBudgetFailure() {
        var area = new OfficeTextFlowRegion(0, 0, 100, 100);
        var obstacles = new[] { new OfficeTextFlowRegion(25, 25, 50, 50) };
        Assert.Throws<InvalidOperationException>(() => OfficeTextFlowRegions.Exclude(area, obstacles,
            () => throw new InvalidOperationException("inspection budget"), default));
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        Assert.Throws<OperationCanceledException>(() => OfficeTextFlowRegions.Exclude(area, obstacles, () => { }, cancelled.Token));
    }
}
