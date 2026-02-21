using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderNoRegexGuardrailTests {
    [Theory]
    [InlineData("OfficeIMO.Reader.Csv")]
    [InlineData("OfficeIMO.Reader.Json")]
    [InlineData("OfficeIMO.Reader.Xml")]
    [InlineData("OfficeIMO.Reader.Text")]
    [InlineData("OfficeIMO.Reader.Html")]
    [InlineData("OfficeIMO.Reader.Zip")]
    [InlineData("OfficeIMO.Reader.Epub")]
    public void ModularReaderAdapters_DoNotUseRegexParsing(string projectFolderName) {
        var projectFolder = Path.Combine(GetRepositoryRoot(), projectFolderName);
        Assert.True(Directory.Exists(projectFolder), "Project folder missing: " + projectFolder);

        var sourceFiles = Directory
            .EnumerateFiles(projectFolder, "*.cs", SearchOption.AllDirectories)
            .Where(static file => !file.Contains(Path.DirectorySeparatorChar + "bin" + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
            .Where(static file => !file.Contains(Path.DirectorySeparatorChar + "obj" + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
            .OrderBy(static file => file, StringComparer.Ordinal)
            .ToArray();

        Assert.NotEmpty(sourceFiles);

        foreach (var file in sourceFiles) {
            var code = File.ReadAllText(file);

            Assert.DoesNotContain("System.Text.RegularExpressions", code, StringComparison.Ordinal);
            Assert.DoesNotContain("Regex.", code, StringComparison.Ordinal);
            Assert.DoesNotContain(" Regex(", code, StringComparison.Ordinal);
            Assert.DoesNotContain("new Regex", code, StringComparison.Ordinal);
        }
    }

    private static string GetRepositoryRoot() {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory != null) {
            if (
                File.Exists(Path.Combine(directory.FullName, "OfficeIMO.sln")) ||
                File.Exists(Path.Combine(directory.FullName, "OfficeImo.sln"))
            ) {
                return directory.FullName;
            }

            directory = directory.Parent;
        }

        throw new DirectoryNotFoundException("Unable to locate OfficeIMO repository root from test runtime base directory.");
    }
}
