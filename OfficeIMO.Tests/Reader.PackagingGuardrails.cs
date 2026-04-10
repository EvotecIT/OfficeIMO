using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderPackagingGuardrailTests {
    [Theory]
    [InlineData("OfficeIMO.Zip/OfficeIMO.Zip.csproj")]
    [InlineData("OfficeIMO.Epub/OfficeIMO.Epub.csproj")]
    [InlineData("OfficeIMO.Reader.Zip/OfficeIMO.Reader.Zip.csproj")]
    [InlineData("OfficeIMO.Reader.Epub/OfficeIMO.Reader.Epub.csproj")]
    [InlineData("OfficeIMO.Reader.Html/OfficeIMO.Reader.Html.csproj")]
    [InlineData("OfficeIMO.Reader.Text/OfficeIMO.Reader.Text.csproj")]
    [InlineData("OfficeIMO.Reader.Csv/OfficeIMO.Reader.Csv.csproj")]
    [InlineData("OfficeIMO.Reader.Json/OfficeIMO.Reader.Json.csproj")]
    [InlineData("OfficeIMO.Reader.Xml/OfficeIMO.Reader.Xml.csproj")]
    public void ModularReaderProjects_RemainPackableAndPublishable(string relativeProjectPath) {
        var projectPath = Path.Combine(GetRepositoryRoot(), relativeProjectPath.Replace('/', Path.DirectorySeparatorChar));
        Assert.True(File.Exists(projectPath), "Project file is missing: " + projectPath);

        var document = XDocument.Load(projectPath);
        var ns = document.Root?.Name.Namespace ?? XNamespace.None;

        var isPackableValues = document.Descendants(ns + "IsPackable")
            .Select(static e => (e.Value ?? string.Empty).Trim())
            .ToArray();
        var isPublishableValues = document.Descendants(ns + "IsPublishable")
            .Select(static e => (e.Value ?? string.Empty).Trim())
            .ToArray();

        Assert.Contains(isPackableValues, static value => string.Equals(value, "true", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(isPublishableValues, static value => string.Equals(value, "true", StringComparison.OrdinalIgnoreCase));
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
