using System;
using System.Linq;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Filters;

namespace OfficeIMO.Benchmarks;

/// <summary>Leaves out BenchmarkDotNet cases of restricted libraries the caller has not enabled.</summary>
internal static class BenchmarkLibraryFilter {
    internal static ManualConfig AddRestrictedLibraryFilter(this ManualConfig config) =>
        config.AddFilter(new SimpleFilter(benchmarkCase => benchmarkCase.Descriptor.Categories.All(category =>
            !category.StartsWith(BenchmarkLibrarySelection.CategoryPrefix, StringComparison.OrdinalIgnoreCase) ||
            BenchmarkLibrarySelection.IsEnabled(category.Substring(BenchmarkLibrarySelection.CategoryPrefix.Length)))));
}
