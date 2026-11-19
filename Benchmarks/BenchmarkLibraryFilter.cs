using System;
using System.Linq;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Filters;

namespace OfficeIMO.Benchmarks;

/// <summary>
/// Leaves out BenchmarkDotNet cases whose <c>library:</c> category names a restricted library the caller has not
/// enabled. Cases of open libraries always run.
/// </summary>
internal static class BenchmarkLibraryFilter {
    internal static ManualConfig AddRestrictedLibraryFilter(this ManualConfig config) =>
        config.AddFilter(new SimpleFilter(benchmarkCase => benchmarkCase.Descriptor.Categories.All(category =>
            !category.StartsWith(BenchmarkLibrarySelection.CategoryPrefix, StringComparison.OrdinalIgnoreCase) ||
            BenchmarkLibrarySelection.IsEnabled(category.Substring(BenchmarkLibrarySelection.CategoryPrefix.Length)))));
}
