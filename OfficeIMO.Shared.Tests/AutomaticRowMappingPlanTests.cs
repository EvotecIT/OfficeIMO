using OfficeIMO.Data;
using System.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class AutomaticRowMappingPlanTests {
    [Fact]
    public void AutomaticMappingCacheRetainsOnlyBoundedHeaderShapes() {
        Assert.True(AutomaticRowMappingPlan<MappedRow>.IsHeaderShapeCacheable(["Value", "Description"]));
        Assert.False(AutomaticRowMappingPlan<MappedRow>.IsHeaderShapeCacheable(
            Enumerable.Repeat("x", AutomaticRowMappingPlan<MappedRow>.MaximumCachedHeaderCount + 1).ToArray()));
        Assert.False(AutomaticRowMappingPlan<MappedRow>.IsHeaderShapeCacheable(
            [new string('x', AutomaticRowMappingPlan<MappedRow>.MaximumCachedHeaderCharacters + 1)]));
    }

    [Fact]
    public void AutomaticMappingBuildsButDoesNotRetainPlansForUncacheableHeaderShapes() {
        string oversizedHeader = new('x', AutomaticRowMappingPlan<MappedRow>.MaximumCachedHeaderCharacters + 1);

        AutomaticRowMappingPlan<MappedRow> first = AutomaticRowMappingPlan<MappedRow>.Create(
            ["Value", oversizedHeader]);
        AutomaticRowMappingPlan<MappedRow> second = AutomaticRowMappingPlan<MappedRow>.Create(
            ["Value", oversizedHeader]);

        Assert.NotSame(first, second);
    }

    private sealed class MappedRow {
        public string? Value { get; set; }
    }
}
