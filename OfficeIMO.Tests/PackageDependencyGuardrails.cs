using System.Text.Json;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PackageDependencyGuardrailTests {
    private const string CurrentMarkdigVersion = "1.3.2";

    [Fact]
    public void MarkdownParityProjects_UseTheSameCurrentMarkdigBaseline() {
        string[] projectPaths = [
            "OfficeIMO.Tests/OfficeIMO.Tests.csproj",
            "OfficeIMO.Markdown.Benchmarks/OfficeIMO.Markdown.Benchmarks.csproj"
        ];

        foreach (var relativeProjectPath in projectPaths) {
            var projectPath = GetRepositoryPath(relativeProjectPath);
            Assert.True(File.Exists(projectPath), "Project file is missing: " + projectPath);

            Assert.Equal(CurrentMarkdigVersion, GetPackageReferenceVersion(projectPath, "Markdig"));
        }
    }

    [Fact]
    public void MarkdownCompatibilityDocs_TrackCurrentMarkdigBaselineVersion() {
        string compatibilityMatrix = File.ReadAllText(GetRepositoryPath("Docs/officeimo.markdown.compatibility-matrix.md"));
        Assert.Contains($"| External comparison package | Markdig `{CurrentMarkdigVersion}`", compatibilityMatrix, StringComparison.Ordinal);

        string competitorRoadmap = File.ReadAllText(GetRepositoryPath("Docs/officeimo.markdown.markdig-competitor-roadmap.md"));
        Assert.Contains($"external parity baseline: Markdig `{CurrentMarkdigVersion}`", competitorRoadmap, StringComparison.Ordinal);

        string correctnessBacklog = File.ReadAllText(GetRepositoryPath("Docs/officeimo.markdown.correctness-backlog.md"));
        Assert.Contains($"`OfficeIMO.Tests` and `OfficeIMO.Markdown.Benchmarks` both reference Markdig `{CurrentMarkdigVersion}`", correctnessBacklog, StringComparison.Ordinal);

        string packageCompatibility = File.ReadAllText(GetRepositoryPath("OfficeIMO.Markdown/COMPATIBILITY.md"));
        Assert.Contains($"curated Markdig {CurrentMarkdigVersion} parity cases", packageCompatibility, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownCompatibilityDocs_TrackCurrentFixtureBaselineCounts() {
        int commonMarkFixtureCount = CountJsonArrayEntries("OfficeIMO.Tests/Markdown/Fixtures/CommonMark/commonmark-0.31.2-smoke.json");
        int gfmFixtureCount = CountJsonArrayEntries("OfficeIMO.Tests/Markdown/Fixtures/GitHubFlavoredMarkdown/cmark-gfm-extensions-smoke.json");

        string compatibilityMatrix = File.ReadAllText(GetRepositoryPath("Docs/officeimo.markdown.compatibility-matrix.md"));
        Assert.Contains($"| CommonMark reference | {commonMarkFixtureCount} of 652 official CommonMark `0.31.2` examples pinned as smoke fixtures |", compatibilityMatrix, StringComparison.Ordinal);
        Assert.Contains($"| GFM reference | {gfmFixtureCount} cmark-gfm extension smoke fixtures plus focused OfficeIMO supplements for upstream ignored-autolink crash and query/fragment autolink regressions |", compatibilityMatrix, StringComparison.Ordinal);

        string competitorRoadmap = File.ReadAllText(GetRepositoryPath("Docs/officeimo.markdown.markdig-competitor-roadmap.md"));
        Assert.Contains($"standards smoke baseline: {commonMarkFixtureCount} CommonMark `0.31.2` fixtures, {gfmFixtureCount} cmark-gfm extension fixtures", competitorRoadmap, StringComparison.Ordinal);

        string parityGapPlan = File.ReadAllText(GetRepositoryPath("Docs/officeimo.markdown.markdig-parity-gap-plan.md"));
        Assert.Contains($"| CommonMark corpus | {commonMarkFixtureCount} of 652 official CommonMark `0.31.2` examples pinned as smoke fixtures |", parityGapPlan, StringComparison.Ordinal);
        Assert.Contains($"| GFM corpus | {gfmFixtureCount} cmark-gfm extension smoke fixtures plus focused crash/regression coverage |", parityGapPlan, StringComparison.Ordinal);

        string packageCompatibility = File.ReadAllText(GetRepositoryPath("OfficeIMO.Markdown/COMPATIBILITY.md"));
        Assert.Contains($"includes {commonMarkFixtureCount} pinned CommonMark 0.31.2 fixtures, {gfmFixtureCount} cmark-gfm smoke fixtures", packageCompatibility, StringComparison.Ordinal);
    }

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
    [InlineData("OfficeIMO.Rtf/OfficeIMO.Rtf.csproj")]
    [InlineData("OfficeIMO.Word.Rtf/OfficeIMO.Word.Rtf.csproj")]
    [InlineData("OfficeIMO.Rtf.Pdf/OfficeIMO.Rtf.Pdf.csproj")]
    [InlineData("OfficeIMO.Drawing/OfficeIMO.Drawing.csproj")]
    [InlineData("OfficeIMO.Pdf/OfficeIMO.Pdf.csproj")]
    [InlineData("OfficeIMO.Word.Pdf/OfficeIMO.Word.Pdf.csproj")]
    [InlineData("OfficeIMO.Excel.Pdf/OfficeIMO.Excel.Pdf.csproj")]
    [InlineData("OfficeIMO.Markdown.Pdf/OfficeIMO.Markdown.Pdf.csproj")]
    [InlineData("OfficeIMO.PowerPoint.Pdf/OfficeIMO.PowerPoint.Pdf.csproj")]
    [InlineData("OfficeIMO.Html.Pdf/OfficeIMO.Html.Pdf.csproj")]
    [InlineData("OfficeIMO.Reader.Pdf/OfficeIMO.Reader.Pdf.csproj")]
    [InlineData("OfficeIMO.Reader.Rtf/OfficeIMO.Reader.Rtf.csproj")]
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

    [Fact]
    public void RtfHtmlBridge_IsUnifiedIntoOfficeIMOHtml() {
        var projectPath = GetRepositoryPath("OfficeIMO.Html/OfficeIMO.Html.csproj");
        Assert.True(File.Exists(projectPath), "Project file is missing: " + projectPath);
        Assert.False(Directory.Exists(GetRepositoryPath("OfficeIMO.Rtf.Html")), "Retired RTF HTML project folder should not be restored.");
        Assert.False(Directory.Exists(GetRepositoryPath("OfficeIMO.Html.Rtf")), "Retired HTML RTF project folder should not be restored.");

        var document = XDocument.Load(projectPath);
        var ns = document.Root?.Name.Namespace ?? XNamespace.None;

        Assert.Equal("OfficeIMO.Html", (string?)document.Descendants(ns + "PackageId").Single());
        Assert.Equal("OfficeIMO.Html", (string?)document.Descendants(ns + "AssemblyName").Single());

        var exportedTypeNames = typeof(OfficeIMO.Html.HtmlToRtfOptions)
            .Assembly
            .GetExportedTypes()
            .Select(static type => type.FullName ?? type.Name)
            .ToArray();

        Assert.Contains("OfficeIMO.Html.HtmlToRtfOptions", exportedTypeNames);
        Assert.Contains("OfficeIMO.Html.RtfToHtmlOptions", exportedTypeNames);
        Assert.DoesNotContain(exportedTypeNames, static typeName => typeName.Contains(".RtfHtml", StringComparison.Ordinal));

        var projectReferences = document
            .Descendants(ns + "ProjectReference")
            .Select(static e => NormalizeProjectPath((string?)e.Attribute("Include")))
            .ToArray();

        Assert.Contains(projectReferences, static include => include.EndsWith("OfficeIMO.Rtf/OfficeIMO.Rtf.csproj", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(projectReferences, static include => include.Contains("OfficeIMO.Rtf.Html", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void RetiredRtfHtmlPackages_AreNotReferencedBySolutionOrProjects() {
        string[] retiredPackageIds = ["OfficeIMO.Rtf.Html", "OfficeIMO.Html.Rtf"];

        var solutionPath = GetRepositoryPath("OfficeIMO.sln");
        Assert.True(File.Exists(solutionPath), "Solution file is missing: " + solutionPath);

        var solutionText = File.ReadAllText(solutionPath);
        foreach (var retiredPackageId in retiredPackageIds) {
            Assert.DoesNotContain(retiredPackageId, solutionText, StringComparison.OrdinalIgnoreCase);
        }

        var projectFiles = EnumerateProjectFiles();
        foreach (var projectFile in projectFiles) {
            var document = XDocument.Load(projectFile);
            var ns = document.Root?.Name.Namespace ?? XNamespace.None;

            var packageIds = document
                .Descendants(ns + "PackageId")
                .Select(static element => (string?)element)
                .Where(static value => !string.IsNullOrWhiteSpace(value))
                .ToArray();

            var packageReferences = document
                .Descendants(ns + "PackageReference")
                .Select(static element => (string?)element.Attribute("Include"))
                .Where(static value => !string.IsNullOrWhiteSpace(value))
                .ToArray();

            var projectReferences = document
                .Descendants(ns + "ProjectReference")
                .Select(static element => NormalizeProjectPath((string?)element.Attribute("Include")))
                .Where(static value => !string.IsNullOrWhiteSpace(value))
                .ToArray();

            foreach (var retiredPackageId in retiredPackageIds) {
                Assert.DoesNotContain(packageIds, value => string.Equals(value, retiredPackageId, StringComparison.OrdinalIgnoreCase));
                Assert.DoesNotContain(packageReferences, value => string.Equals(value, retiredPackageId, StringComparison.OrdinalIgnoreCase));
                Assert.DoesNotContain(projectReferences, value => value.Contains(retiredPackageId, StringComparison.OrdinalIgnoreCase));
            }
        }

        var projectBuildPath = GetRepositoryPath("Build/project.build.json");
        Assert.True(File.Exists(projectBuildPath), "Project build file is missing: " + projectBuildPath);

        var projectBuildText = File.ReadAllText(projectBuildPath);
        foreach (var retiredPackageId in retiredPackageIds) {
            Assert.DoesNotContain(retiredPackageId, projectBuildText, StringComparison.OrdinalIgnoreCase);
        }
    }

    [Fact]
    public void RetiredRtfHtmlNamespaces_AreNotUsedBySourceFiles() {
        string[] retiredNamespaces = ["OfficeIMO.Rtf.Html", "OfficeIMO.Html.Rtf"];

        var sourceFiles = Directory.EnumerateFiles(GetRepositoryRoot(), "*.cs", SearchOption.AllDirectories)
            .Where(static path => !path.Contains($"{Path.DirectorySeparatorChar}bin{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase))
            .Where(static path => !path.Contains($"{Path.DirectorySeparatorChar}obj{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase))
            .Where(static path => !path.Contains($"{Path.DirectorySeparatorChar}Ignore{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase))
            .Where(static path => new FileInfo(path).Length > 0)
            .ToArray();

        foreach (var sourceFile in sourceFiles) {
            string source = File.ReadAllText(sourceFile);
            foreach (var retiredNamespace in retiredNamespaces) {
                Assert.DoesNotContain($"namespace {retiredNamespace}", source, StringComparison.Ordinal);
                Assert.DoesNotContain($"using {retiredNamespace}", source, StringComparison.Ordinal);
            }
        }
    }

    [Fact]
    public void RtfPackages_AreIncludedInProjectBuildVersionMap() {
        var projectBuildPath = GetRepositoryPath("Build/project.build.json");
        Assert.True(File.Exists(projectBuildPath), "Project build file is missing: " + projectBuildPath);

        using JsonDocument document = JsonDocument.Parse(File.ReadAllText(projectBuildPath));
        JsonElement expectedVersionMap = document.RootElement.GetProperty("ExpectedVersionMap");

        Assert.Equal("0.1.X", expectedVersionMap.GetProperty("OfficeIMO.Rtf").GetString());
        Assert.Equal("0.1.X", expectedVersionMap.GetProperty("OfficeIMO.Word.Rtf").GetString());
        Assert.Equal("0.1.X", expectedVersionMap.GetProperty("OfficeIMO.Rtf.Pdf").GetString());
        Assert.Equal("0.0.X", expectedVersionMap.GetProperty("OfficeIMO.Reader.Rtf").GetString());
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

    private static string[] EnumerateProjectFiles() =>
        Directory.EnumerateFiles(GetRepositoryRoot(), "*.csproj", SearchOption.AllDirectories)
            .Where(static path => !path.Contains($"{Path.DirectorySeparatorChar}bin{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase))
            .Where(static path => !path.Contains($"{Path.DirectorySeparatorChar}obj{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase))
            .Where(static path => !path.Contains($"{Path.DirectorySeparatorChar}Ignore{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase))
            .Where(static path => new FileInfo(path).Length > 0)
            .ToArray();

    private static int CountJsonArrayEntries(string relativePath) {
        var path = GetRepositoryPath(relativePath);
        Assert.True(File.Exists(path), "Fixture file is missing: " + path);

        using JsonDocument document = JsonDocument.Parse(File.ReadAllText(path));
        Assert.Equal(JsonValueKind.Array, document.RootElement.ValueKind);
        return document.RootElement.GetArrayLength();
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

    private static string? GetPackageReferenceVersion(string projectPath, string packageId) {
        var document = XDocument.Load(projectPath);
        var ns = document.Root?.Name.Namespace ?? XNamespace.None;

        return document
            .Descendants(ns + "PackageReference")
            .Where(element => string.Equals((string?)element.Attribute("Include"), packageId, StringComparison.OrdinalIgnoreCase))
            .Select(static element => (string?)element.Attribute("Version"))
            .SingleOrDefault();
    }
}
