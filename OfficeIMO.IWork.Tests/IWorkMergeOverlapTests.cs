using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;
using System.Threading;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkMergeOverlapTests {
    [Fact]
    public void Every_nested_or_intersecting_rectangle_is_identified_without_rejecting_adjacent_ranges() {
        IWorkTableMergeRange[] ranges = {
            new(1, 1, 8, 8), new(3, 3, 4, 4), new(9, 1, 10, 2),
            new(9, 3, 10, 4), new(1, 9, 8, 10), new(2, 9, 2, 10), new(8, 8, 9, 9)
        };
        Assert.Equal(new[] { 0, 1, 4, 5, 6 }, IWorkMergeOverlapIndex.FindOverlapIndexes(ranges, 10, default));
    }

    [Fact]
    public void Conflict_inventory_matches_independent_rectangle_intersection_oracle() {
        var random = new Random(1451);
        for (int sample = 0; sample < 8; sample++) {
            IWorkTableMergeRange[] ranges = Enumerable.Range(0, 128).Select(_ => {
                int row = random.Next(1, 129), column = random.Next(1, 17);
                return new IWorkTableMergeRange(row, column, row + random.Next(0, 4), Math.Min(16, column + random.Next(0, 3)));
            }).ToArray();
            int[] expected = Enumerable.Range(0, ranges.Length).Where(index => ranges.Where((_, other) => other != index)
                .Any(other => ranges[index].FirstRow <= other.LastRow && other.FirstRow <= ranges[index].LastRow
                    && ranges[index].FirstColumn <= other.LastColumn && other.FirstColumn <= ranges[index].LastColumn)).ToArray();
            Assert.Equal(expected, IWorkMergeOverlapIndex.FindOverlapIndexes(ranges, 16, default));
        }
    }

    [Fact]
    public void Large_dense_and_sparse_sets_have_complete_conflict_inventories() {
        IWorkTableMergeRange[] dense = Enumerable.Range(0, 20_000).Select(_ => new IWorkTableMergeRange(1, 1, 4, 8)).ToArray();
        Assert.Equal(Enumerable.Range(0, dense.Length), IWorkMergeOverlapIndex.FindOverlapIndexes(dense, 8, default));
        IWorkTableMergeRange[] sparse = Enumerable.Range(0, 20_000).Select(index => new IWorkTableMergeRange(index * 3 + 1, 1, index * 3 + 2, 2)).ToArray();
        Assert.Empty(IWorkMergeOverlapIndex.FindOverlapIndexes(sparse, 2, default));
    }

    [Fact]
    public void Canceled_conflict_assessment_does_not_sort_or_allocate_an_index() {
        Assert.Throws<OperationCanceledException>(() => IWorkMergeOverlapIndex.FindOverlapIndexes(
            new[] { new IWorkTableMergeRange(1, 1, 2, 2), new IWorkTableMergeRange(1, 1, 3, 3) }, 3, new CancellationToken(true)));
    }
}
