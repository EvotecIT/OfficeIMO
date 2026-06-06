using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PackageDependencyGuardrailTests {
    [Fact]
    public void Projects_DoNotReferenceImageSharpPackage() {
        var projectFiles = Directory.EnumerateFiles(GetRepositoryRoot(), "*.csproj", SearchOption.AllDirectories)
            .Where(static path => !path.Contains($"{Path.DirectorySeparatorChar}bin{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase))
            .Where(static path => !path.Contains($"{Path.DirectorySeparatorChar}obj{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase))
            .Where(static path => !path.Contains($"{Path.DirectorySeparatorChar}Ignore{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase))
            .Where(static path => new FileInfo(path).Length > 0)
            .ToArray();

        var offenders = projectFiles
            .Where(ProjectReferencesImageSharp)
            .ToArray();

        Assert.Empty(offenders);
    }

    [Fact]
    public void Projects_DoNotReferenceSixLaborsFontsPackage() {
        var projectFiles = Directory.EnumerateFiles(GetRepositoryRoot(), "*.csproj", SearchOption.AllDirectories)
            .Where(static path => !path.Contains($"{Path.DirectorySeparatorChar}bin{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase))
            .Where(static path => !path.Contains($"{Path.DirectorySeparatorChar}obj{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase))
            .Where(static path => !path.Contains($"{Path.DirectorySeparatorChar}Ignore{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase))
            .Where(static path => new FileInfo(path).Length > 0)
            .ToArray();

        var offenders = projectFiles
            .Where(ProjectReferencesSixLaborsFonts)
            .ToArray();

        Assert.Empty(offenders);
    }

    [Theory]
    [InlineData("OfficeIMO.Drawing/OfficeIMO.Drawing.csproj")]
    [InlineData("OfficeIMO.Pdf/OfficeIMO.Pdf.csproj")]
    [InlineData("OfficeIMO.Word.Pdf/OfficeIMO.Word.Pdf.csproj")]
    [InlineData("OfficeIMO.Excel.Pdf/OfficeIMO.Excel.Pdf.csproj")]
    [InlineData("OfficeIMO.Markdown.Pdf/OfficeIMO.Markdown.Pdf.csproj")]
    [InlineData("OfficeIMO.PowerPoint.Pdf/OfficeIMO.PowerPoint.Pdf.csproj")]
    [InlineData("OfficeIMO.Html.Pdf/OfficeIMO.Html.Pdf.csproj")]
    [InlineData("OfficeIMO.Reader.Pdf/OfficeIMO.Reader.Pdf.csproj")]
    public void DependencyLightProjects_HaveNoPackageReferences(string relativeProjectPath) {
        var projectPath = GetRepositoryPath(relativeProjectPath);
        Assert.True(File.Exists(projectPath), "Project file is missing: " + projectPath);

        var document = XDocument.Load(projectPath);
        var ns = document.Root?.Name.Namespace ?? XNamespace.None;

        var references = document
            .Descendants(ns + "PackageReference")
            .Select(static e => (string?)e.Attribute("Include") ?? string.Empty)
            .Where(static include => !string.IsNullOrWhiteSpace(include))
            .ToArray();

        Assert.Empty(references);
    }

    [Theory]
    [InlineData("OfficeIMO.Word/OfficeIMO.Word.csproj")]
    [InlineData("OfficeIMO.Excel/OfficeIMO.Excel.csproj")]
    [InlineData("OfficeIMO.Visio/OfficeIMO.Visio.csproj")]
    [InlineData("OfficeIMO.Word.Html/OfficeIMO.Word.Html.csproj")]
    [InlineData("OfficeIMO.Word.Markdown/OfficeIMO.Word.Markdown.csproj")]
    public void ImageAndColorConsumers_ReferenceOfficeImoDrawing(string relativeProjectPath) {
        var projectPath = GetRepositoryPath(relativeProjectPath);
        Assert.True(File.Exists(projectPath), "Project file is missing: " + projectPath);

        var document = XDocument.Load(projectPath);
        var ns = document.Root?.Name.Namespace ?? XNamespace.None;

        var references = document
            .Descendants(ns + "ProjectReference")
            .Select(static e => NormalizeProjectPath((string?)e.Attribute("Include")))
            .Where(static include => include.EndsWith("OfficeIMO.Drawing/OfficeIMO.Drawing.csproj", StringComparison.OrdinalIgnoreCase))
            .ToArray();

        Assert.Single(references);
    }

    [Theory]
    [InlineData("OfficeIMO.Reader/OfficeIMO.Reader.csproj")]
    [InlineData("OfficeIMO.Reader.Json/OfficeIMO.Reader.Json.csproj")]
    [InlineData("OfficeIMO.MarkdownRenderer/OfficeIMO.MarkdownRenderer.csproj")]
    [InlineData("OfficeIMO.MarkdownRenderer.SamplePlugin/OfficeIMO.MarkdownRenderer.SamplePlugin.csproj")]
    [InlineData("OfficeIMO.GoogleWorkspace/OfficeIMO.GoogleWorkspace.csproj")]
    [InlineData("OfficeIMO.Excel.GoogleSheets/OfficeIMO.Excel.GoogleSheets.csproj")]
    [InlineData("OfficeIMO.Word.GoogleDocs/OfficeIMO.Word.GoogleDocs.csproj")]
    public void SystemTextJsonPackageReference_IsLimitedToNonInboxTargets(string relativeProjectPath) {
        var projectPath = GetRepositoryPath(relativeProjectPath);
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
        var projectPath = GetRepositoryPath(relativeProjectPath);
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

    private static string GetRepositoryPath(string relativePath) {
        Assert.False(Path.IsPathRooted(relativePath), "Repository-relative path must not be rooted: " + relativePath);

        var repositoryRoot = Path.GetFullPath(GetRepositoryRoot());
        if (!repositoryRoot.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal)) {
            repositoryRoot += Path.DirectorySeparatorChar;
        }

        var parts = NormalizeProjectPath(relativePath)
            .Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
        var combinedPath = repositoryRoot;
        foreach (var part in parts) {
            Assert.False(Path.IsPathRooted(part), "Repository-relative path segment must not be rooted: " + relativePath);
            combinedPath = AppendRepositoryPathSegment(combinedPath, part);
        }

        combinedPath = Path.GetFullPath(combinedPath);

        Assert.True(
            combinedPath.StartsWith(repositoryRoot, StringComparison.Ordinal),
            "Repository-relative path must stay under repository root: " + relativePath);
        return combinedPath;
    }

    private static string AppendRepositoryPathSegment(string basePath, string segment) =>
        basePath.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal)
            ? basePath + segment
            : basePath + Path.DirectorySeparatorChar + segment;

    private static string NormalizeProjectPath(string? path) =>
        (path ?? string.Empty).Replace('\\', '/');

    private static bool ProjectReferencesImageSharp(string projectPath) {
        var document = XDocument.Load(projectPath);
        var ns = document.Root?.Name.Namespace ?? XNamespace.None;

        return document
            .Descendants(ns + "PackageReference")
            .Any(static e => string.Equals((string?)e.Attribute("Include"), "SixLabors.ImageSharp", StringComparison.Ordinal));
    }

    private static bool ProjectReferencesSixLaborsFonts(string projectPath) {
        var document = XDocument.Load(projectPath);
        var ns = document.Root?.Name.Namespace ?? XNamespace.None;

        return document
            .Descendants(ns + "PackageReference")
            .Any(static e => string.Equals((string?)e.Attribute("Include"), "SixLabors.Fonts", StringComparison.Ordinal));
    }
}
