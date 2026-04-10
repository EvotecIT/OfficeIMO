using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PackageDependencyGuardrailTests {
    [Theory]
    [InlineData("OfficeIMO.Reader/OfficeIMO.Reader.csproj")]
    [InlineData("OfficeIMO.Reader.Json/OfficeIMO.Reader.Json.csproj")]
    [InlineData("OfficeIMO.MarkdownRenderer/OfficeIMO.MarkdownRenderer.csproj")]
    [InlineData("OfficeIMO.MarkdownRenderer.SamplePlugin/OfficeIMO.MarkdownRenderer.SamplePlugin.csproj")]
    [InlineData("OfficeIMO.GoogleWorkspace/OfficeIMO.GoogleWorkspace.csproj")]
    [InlineData("OfficeIMO.Excel.GoogleSheets/OfficeIMO.Excel.GoogleSheets.csproj")]
    [InlineData("OfficeIMO.Word.GoogleDocs/OfficeIMO.Word.GoogleDocs.csproj")]
    public void SystemTextJsonPackageReference_IsLimitedToNonInboxTargets(string relativeProjectPath) {
        var projectPath = Path.Combine(GetRepositoryRoot(), relativeProjectPath.Replace('/', Path.DirectorySeparatorChar));
        Assert.True(File.Exists(projectPath), "Project file is missing: " + projectPath);

        var document = XDocument.Load(projectPath);
        var ns = document.Root?.Name.Namespace ?? XNamespace.None;

        var references = document
            .Descendants(ns + "PackageReference")
            .Where(static e => string.Equals((string?)e.Attribute("Include"), "System.Text.Json", StringComparison.Ordinal))
            .ToArray();

        Assert.Single(references);

        var parentItemGroup = references[0].Parent;
        Assert.NotNull(parentItemGroup);

        var condition = (string?)parentItemGroup!.Attribute("Condition");
        Assert.False(string.IsNullOrWhiteSpace(condition));
        Assert.Contains("netstandard2.0", condition!, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("net472", condition!, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("net8.0", condition!, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("net10.0", condition!, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("OfficeIMO.CSV/OfficeIMO.CSV.csproj")]
    [InlineData("OfficeIMO.CSV.Tests/OfficeIMO.CSV.Tests.csproj")]
    public void NetFrameworkReferenceAssemblies_AreLimitedToNet472(string relativeProjectPath) {
        var projectPath = Path.Combine(GetRepositoryRoot(), relativeProjectPath.Replace('/', Path.DirectorySeparatorChar));
        Assert.True(File.Exists(projectPath), "Project file is missing: " + projectPath);

        var document = XDocument.Load(projectPath);
        var ns = document.Root?.Name.Namespace ?? XNamespace.None;

        var references = document
            .Descendants(ns + "PackageReference")
            .Where(static e => string.Equals((string?)e.Attribute("Include"), "Microsoft.NETFramework.ReferenceAssemblies", StringComparison.Ordinal))
            .ToArray();

        Assert.Single(references);

        var parentItemGroup = references[0].Parent;
        Assert.NotNull(parentItemGroup);

        var condition = (string?)parentItemGroup!.Attribute("Condition");
        Assert.False(string.IsNullOrWhiteSpace(condition));
        Assert.Contains("net472", condition!, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("net8.0", condition!, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("net10.0", condition!, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("netstandard2.0", condition!, StringComparison.OrdinalIgnoreCase);
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
