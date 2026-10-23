using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class DrawingTextSpacingAllocatorTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void EarlierPriorityIsExhaustedBeforeProportionalPartialPriority(bool compress) {
        OfficeTextSpacingResource[] resources = {
            new(2, 10, 3D, 6D),
            new(0, 0, 1D, 2D),
            new(1, 10, 1D, 2D),
            new(0, 10, 2D, 4D)
        };

        OfficeTextSpacingAllocation result = OfficeTextSpacingAllocator.Allocate(3, resources, compress ? -4D : 8D);

        double scale = compress ? -0.5D : 1D;
        Assert.Equal(new[] { 4D * scale, 1D * scale, 3D * scale }, result.Deltas);
        Assert.Equal(compress ? -4D : 8D, result.AppliedAdjustment);
        Assert.Equal(0D, result.UnfulfilledAdjustment);
    }

    [Fact]
    public void PartialPriorityUsesCapacityAndCarriesTheRepresentableRemainder() {
        OfficeTextSpacingResource[] resources = { new(2, 0, 1D, 1D), new(0, 0, 1D, 1D), new(1, 0, 1D, 1D) };

        OfficeTextSpacingAllocation result = OfficeTextSpacingAllocator.Allocate(3, resources, 1D);

        Assert.Equal(1D / 3D, result.Deltas[2]);
        Assert.All(result.Deltas, delta => Assert.InRange(delta, 0.333333333333333D, 0.333333333333334D));
        Assert.Equal(1D, result.Deltas.Sum());
        Assert.Equal(1D, result.AppliedAdjustment);
        Assert.Equal(0D, result.UnfulfilledAdjustment);
        Assert.Equal(result.Deltas, OfficeTextSpacingAllocator.Allocate(3, resources, 1D).Deltas);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExhaustedCapacityReportsTheSignedUnfulfilledRequest(bool compress) {
        OfficeTextSpacingResource[] resources = { new(0, 0, 2D, 2D), new(1, 1, 1D, 1D) };

        OfficeTextSpacingAllocation result = OfficeTextSpacingAllocator.Allocate(2, resources, compress ? -5D : 5D);

        double sign = compress ? -1D : 1D;
        Assert.Equal(new[] { 2D * sign, sign }, result.Deltas);
        Assert.Equal(3D * sign, result.AppliedAdjustment);
        Assert.Equal(2D * sign, result.UnfulfilledAdjustment);
    }

    [Fact]
    public void EmptyAndZeroCapacityRequestsRetainTheirActualResidual() {
        OfficeTextSpacingAllocation empty = OfficeTextSpacingAllocator.Allocate(0, Array.Empty<OfficeTextSpacingResource>(), -3D);
        Assert.Empty(empty.Deltas);
        Assert.Equal(0D, empty.AppliedAdjustment);
        Assert.Equal(-3D, empty.UnfulfilledAdjustment);

        OfficeTextSpacingResource[] resources = { new(1, 0, 0D, 0D) };
        OfficeTextSpacingAllocation zero = OfficeTextSpacingAllocator.Allocate(3, resources, 0D);
        Assert.Equal(new[] { 0D, 0D, 0D }, zero.Deltas);
        Assert.Equal(0D, zero.AppliedAdjustment);
        Assert.Equal(0D, zero.UnfulfilledAdjustment);
        Assert.Equal(4D, OfficeTextSpacingAllocator.Allocate(3, resources, 4D).UnfulfilledAdjustment);
    }

    [Fact]
    public void OneOpportunityAccumulatesDifferentPrioritiesAndTheResultCannotBeMutated() {
        OfficeTextSpacingResource[] resources = { new(0, 0, 1D, 1D), new(0, 1, 1D, 1D), new(1, 1, 2D, 2D) };
        OfficeTextSpacingAllocation result = OfficeTextSpacingAllocator.Allocate(2, resources, 3D);

        Assert.Equal(5D / 3D, result.Deltas[0], 14);
        Assert.Equal(4D / 3D, result.Deltas[1], 14);
        Assert.Equal(result.Deltas.Sum(), result.AppliedAdjustment);
        Assert.Equal(3D - result.AppliedAdjustment, result.UnfulfilledAdjustment);
        double[] snapshot = result.Deltas.ToArray();
        resources[0] = new OfficeTextSpacingResource(1, 0, 10D, 10D);
        Assert.Equal(snapshot, result.Deltas);
        IList<double> exposed = Assert.IsAssignableFrom<IList<double>>(result.Deltas);
        Assert.Throws<NotSupportedException>(() => exposed[0] = 100D);
    }

    [Fact]
    public void SubnormalCallerUnitsAreCarriedWithoutInventingAnAllocationQuantum() {
        OfficeTextSpacingResource[] resources = { new(0, 0, double.Epsilon, double.Epsilon), new(1, 0, double.Epsilon, double.Epsilon) };

        OfficeTextSpacingAllocation result = OfficeTextSpacingAllocator.Allocate(2, resources, double.Epsilon);

        Assert.Equal(new[] { 0D, double.Epsilon }, result.Deltas);
        Assert.Equal(double.Epsilon, result.AppliedAdjustment);
        Assert.Equal(0D, result.UnfulfilledAdjustment);
    }

    [Fact]
    public void FiniteLargeResourcesDoNotOverflowWhileComputingAPartialRatio() {
        OfficeTextSpacingResource[] resources = { new(0, 0, 0D, double.MaxValue / 2D), new(1, 0, 0D, double.MaxValue / 2D) };

        OfficeTextSpacingAllocation result = OfficeTextSpacingAllocator.Allocate(2, resources, double.MaxValue / 2D);

        Assert.Equal(new[] { double.MaxValue / 4D, double.MaxValue / 4D }, result.Deltas);
        Assert.Equal(double.MaxValue / 2D, result.AppliedAdjustment);
        Assert.Equal(0D, result.UnfulfilledAdjustment);
    }

    [Fact]
    public void InvalidCapacitiesRequestsAndOpportunityMappingsAreRejected() {
        foreach (double invalid in new[] { -1D, double.NaN, double.NegativeInfinity, double.PositiveInfinity }) {
            Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeTextSpacingResource(0, 0, invalid, 0D));
            Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeTextSpacingResource(0, 0, 0D, invalid));
        }
        foreach (double invalid in new[] { double.NaN, double.NegativeInfinity, double.PositiveInfinity })
            Assert.Throws<ArgumentOutOfRangeException>(() => OfficeTextSpacingAllocator.Allocate(0, Array.Empty<OfficeTextSpacingResource>(), invalid));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeTextSpacingResource(-1, 0, 0D, 0D));
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeTextSpacingAllocator.Allocate(-1, Array.Empty<OfficeTextSpacingResource>(), 0D));
        Assert.Throws<ArgumentNullException>(() => OfficeTextSpacingAllocator.Allocate(1, null!, 0D));
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeTextSpacingAllocator.Allocate(1, new[] { new OfficeTextSpacingResource(1, 0, 0D, 0D) }, 0D));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AggregateOverflowIsRejectedEvenWhenTheRequestIsZero(bool compressionCapacity) {
        OfficeTextSpacingResource[] resources = compressionCapacity
            ? new[] { new OfficeTextSpacingResource(0, 0, double.MaxValue, 0D), new OfficeTextSpacingResource(1, 1, double.MaxValue, 0D) }
            : new[] { new OfficeTextSpacingResource(0, 0, 0D, double.MaxValue), new OfficeTextSpacingResource(1, 1, 0D, double.MaxValue) };

        Assert.Throws<OverflowException>(() => OfficeTextSpacingAllocator.Allocate(2, resources, 0D));
    }
}
