using System;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingRasterResamplingPlanningTests {
    [Fact]
    public void PlanningIncludesSamplingScratchAndIdenticalDimensionsUseTheCopyPath() {
        long simple = OfficeRasterResampler.GetResizeWorkingSetBytes(9, 7, 3, 2, OfficeRasterResamplingMode.Bilinear);
        long cubic = OfficeRasterResampler.GetResizeWorkingSetBytes(9, 7, 3, 2, OfficeRasterResamplingMode.Bicubic);
        Assert.True(simple > (9 * 7 + 3 * 2) * 4);
        Assert.True(cubic > simple);
        Assert.Equal(
            OfficeRasterResampler.GetResizeWorkingSetBytes(9, 7, 9, 7, OfficeRasterResamplingMode.Bilinear),
            OfficeRasterResampler.GetResizeWorkingSetBytes(9, 7, 9, 7, OfficeRasterResamplingMode.Bicubic));
    }

    [Fact]
    public void PlanningRejectsScratchThatCannotFitWithoutAllocatingTheLargeImage() {
        Assert.Throws<ArgumentException>(() => OfficeRasterResampler.GetResizeWorkingSetBytes(9000, 5000, 4000, 3000, OfficeRasterResamplingMode.Bicubic));
    }

    [Fact]
    public void PlanningObservesCancellationBeforeSamplingWork() {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => OfficeRasterResampler.GetResizeWorkingSetBytes(9, 7, 3, 2,
            OfficeRasterResamplingMode.Bicubic, cancellation.Token));
    }
}
