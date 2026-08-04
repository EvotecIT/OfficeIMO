using Xunit;

namespace OfficeIMO.Tests;

[CollectionDefinition(Name, DisableParallelization = true)]
public sealed class PowerPointNonParallelCollection {
    public const string Name = "PowerPoint non-parallel integration";
}
