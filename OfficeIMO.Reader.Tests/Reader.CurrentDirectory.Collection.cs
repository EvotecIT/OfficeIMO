using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Isolates fixtures that change the process working directory from every other test collection.</summary>
[CollectionDefinition(Name, DisableParallelization = true)]
public sealed class ReaderCurrentDirectoryCollection {
    public const string Name = "Reader current directory";
}
