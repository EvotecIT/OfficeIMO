using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class DocumentationContracts {
    [Theory]
    [InlineData("Docs/Examples/ExamplesAddingTablesWithStylesAndBorders/README.MD")]
    [InlineData("Website/data/xrefmap.json")]
    public void CurrentDocumentationDoesNotAdvertiseRemovedLaunchOnSaveOptions(string relativePath) {
        string documentation = File.ReadAllText(Path.Combine(GetRepositoryRoot(), relativePath));

        Assert.DoesNotContain("OpenAfterSave", documentation, StringComparison.Ordinal);
    }

    private static string GetRepositoryRoot() {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null) {
            if (File.Exists(Path.Combine(directory.FullName, "OfficeIMO.sln"))) {
                return directory.FullName;
            }
            directory = directory.Parent;
        }

        throw new DirectoryNotFoundException("Could not locate the OfficeIMO repository root.");
    }
}
