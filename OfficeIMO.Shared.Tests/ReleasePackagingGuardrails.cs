using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class ReleasePackagingGuardrails {
    [Fact]
    public void ReadmeInventory_MatchesReleaseMapAndLinkedProjectCatalog() {
        string repositoryRoot = GetRepositoryRoot();
        string readme = File.ReadAllText(Path.Combine(repositoryRoot, "README.md"));
        using JsonDocument buildDocument = JsonDocument.Parse(
            File.ReadAllText(Path.Combine(repositoryRoot, "Build", "project.build.json")));
        int releasePackageCount = buildDocument.RootElement
            .GetProperty("ExpectedVersionMap")
            .EnumerateObject()
            .Count();

        MatchCollection projectHeadings = Regex.Matches(
            readme,
            @"^#### \[(?<name>OfficeIMO\.[^\]]+)\]\((?<path>[^)]+)\)$",
            RegexOptions.Multiline | RegexOptions.CultureInvariant);
        string[] duplicateNames = projectHeadings
            .Select(static match => match.Groups["name"].Value)
            .GroupBy(static name => name, StringComparer.OrdinalIgnoreCase)
            .Where(static group => group.Count() > 1)
            .Select(static group => group.Key)
            .ToArray();

        Assert.Empty(duplicateNames);
        Assert.All(projectHeadings, match =>
            Assert.True(
                File.Exists(Path.Combine(repositoryRoot, match.Groups["path"].Value)),
                "README project link is missing: " + match.Value));
        Assert.Equal(24, CountProjectHeadings(readme, "Native formats and shared foundations"));
        Assert.Equal(26, CountProjectHeadings(readme, "Conversion and cloud bridges"));
        Assert.Equal(28, CountProjectHeadings(readme, "Unified Reader family"));
        Assert.Equal(11, CountProjectHeadings(readme, "Markdown rendering and OfficeIMO Markup"));
        Assert.Equal(89, projectHeadings.Count);

        Assert.Contains($"| Coordinated `3.0.x` release packages | {releasePackageCount} |", readme, StringComparison.Ordinal);
        Assert.Contains($"| Documented package, tool, and example projects below | {projectHeadings.Count} |", readme, StringComparison.Ordinal);
        Assert.Contains("| Native format, foundation, and shared-service packages | 24 |", readme, StringComparison.Ordinal);
        Assert.Contains("| Conversion and cloud bridge packages | 26 |", readme, StringComparison.Ordinal);
        Assert.Contains("| Unified Reader packages and tool | 28 |", readme, StringComparison.Ordinal);
        Assert.Contains("| Markdown renderer and OfficeIMO Markup surfaces | 11 |", readme, StringComparison.Ordinal);
    }

    [Fact]
    public void PackageLocks_DoNotRetainOlderOfficeIMOReleaseLines() {
        string repositoryRoot = GetRepositoryRoot();
        string[] lockFiles = Directory
            .EnumerateFiles(repositoryRoot, "packages.lock.json", SearchOption.AllDirectories)
            .Where(static path => !ContainsBuildOutput(path))
            .ToArray();
        Assert.NotEmpty(lockFiles);

        var staleDependencies = new List<string>();
        foreach (string lockFile in lockFiles) {
            string content = File.ReadAllText(lockFile);
            foreach (Match match in Regex.Matches(
                content,
                "\"OfficeIMO\\.[^\"]+\"\\s*:\\s*\"\\[(?<version>\\d+\\.\\d+\\.\\d+),",
                RegexOptions.CultureInvariant | RegexOptions.IgnoreCase)) {
                if (!string.Equals(match.Groups["version"].Value, "3.0.0", StringComparison.Ordinal)) {
                    staleDependencies.Add(
                        Path.GetRelativePath(repositoryRoot, lockFile)
                        + " -> "
                        + match.Value);
                }
            }
        }

        Assert.Empty(staleDependencies);
    }

    [Fact]
    public void ProjectBuild_IncludesEveryPublishablePackageExactlyOnceAndUsesOneVersion() {
        string repositoryRoot = GetRepositoryRoot();
        string projectBuildPath = Path.Combine(repositoryRoot, "Build", "project.build.json");
        using JsonDocument buildDocument = JsonDocument.Parse(File.ReadAllText(projectBuildPath));

        JsonElement buildRoot = buildDocument.RootElement;
        Dictionary<string, string> expectedVersions = buildRoot
            .GetProperty("ExpectedVersionMap")
            .EnumerateObject()
            .ToDictionary(
                static property => property.Name,
                static property => property.Value.GetString() ?? string.Empty,
                StringComparer.OrdinalIgnoreCase);
        string releaseBand = buildRoot.GetProperty("ExpectedVersion").GetString()
            ?? throw new InvalidDataException("Build/project.build.json must declare ExpectedVersion.");
        Assert.Matches(@"^\d+\.\d+\.X$", releaseBand);
        Assert.NotEmpty(expectedVersions);
        Assert.All(expectedVersions, entry => Assert.Equal(releaseBand, entry.Value));
        HashSet<string> excludedProjects = buildRoot
            .GetProperty("ExcludeProjects")
            .EnumerateArray()
            .Select(static element => element.GetString())
            .Where(static value => !string.IsNullOrWhiteSpace(value))
            .Select(static value => value!)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        Assert.Empty(excludedProjects);

        PackageProject[] packageProjects = Directory
            .EnumerateFiles(repositoryRoot, "*.csproj", SearchOption.AllDirectories)
            .Where(static path => !ContainsBuildOutput(path))
            .Select(ReadPackageProject)
            .Where(static project => project is not null)
            .Select(static project => project!)
            .ToArray();

        string[] duplicatePackageIds = packageProjects
            .GroupBy(static project => project.PackageId, StringComparer.OrdinalIgnoreCase)
            .Where(static group => group.Count() > 1)
            .Select(static group => group.Key)
            .OrderBy(static packageId => packageId, StringComparer.OrdinalIgnoreCase)
            .ToArray();
        Assert.Empty(duplicatePackageIds);

        string[] missingFromBuild = packageProjects
            .Where(project => !expectedVersions.ContainsKey(project.PackageId))
            .Select(static project => project.PackageId)
            .OrderBy(static packageId => packageId, StringComparer.OrdinalIgnoreCase)
            .ToArray();
        Assert.Empty(missingFromBuild);

        string[] staleBuildEntries = expectedVersions.Keys
            .Where(packageId => !packageProjects.Any(project =>
                string.Equals(project.PackageId, packageId, StringComparison.OrdinalIgnoreCase)))
            .OrderBy(static packageId => packageId, StringComparer.OrdinalIgnoreCase)
            .ToArray();
        Assert.Empty(staleBuildEntries);

        PackageProject[] includedProjects = packageProjects
            .Where(project => expectedVersions.ContainsKey(project.PackageId))
            .ToArray();
        foreach (PackageProject project in includedProjects) {
            AssertVersionMatchesReleaseBand(project, expectedVersions[project.PackageId]);
        }

        string[] releaseVersions = includedProjects
            .Select(static project => project.Version)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        Assert.Single(releaseVersions);
    }

    [Fact]
    public void SolutionReleaseConfiguration_IncludesEveryPublishablePackage() {
        string repositoryRoot = GetRepositoryRoot();
        string solution = File.ReadAllText(Path.Combine(repositoryRoot, "OfficeIMO.sln"));
        PackageProject[] packageProjects = Directory
            .EnumerateFiles(repositoryRoot, "*.csproj", SearchOption.AllDirectories)
            .Where(static path => !ContainsBuildOutput(path))
            .Select(ReadPackageProject)
            .Where(static project => project is not null)
            .Select(static project => project!)
            .ToArray();

        Assert.All(packageProjects, project => {
            Match projectDeclaration = Assert.Single(Regex.Matches(
                solution,
                $@"^Project\(""[^""]+""\) = ""{Regex.Escape(project.ProjectName)}"", ""[^""]+"", ""\{{(?<guid>[A-F0-9-]+)\}}""\r?$",
                RegexOptions.Multiline | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase)
                .Cast<Match>());
            string projectGuid = projectDeclaration.Groups["guid"].Value;

            Assert.Contains(
                $"{{{projectGuid}}}.Release|Any CPU.Build.0 = Release|Any CPU",
                solution,
                StringComparison.OrdinalIgnoreCase);
        });
    }

    private static PackageProject? ReadPackageProject(string projectPath) {
        XDocument document = XDocument.Load(projectPath);
        XNamespace ns = document.Root?.Name.Namespace ?? XNamespace.None;
        string? packageId = document.Descendants(ns + "PackageId").Select(static element => element.Value).FirstOrDefault();
        string? version = document.Descendants(ns + "VersionPrefix").Select(static element => element.Value).FirstOrDefault();
        if (string.IsNullOrWhiteSpace(packageId) || string.IsNullOrWhiteSpace(version)) {
            return null;
        }

        bool isPackable = !document
            .Descendants(ns + "IsPackable")
            .Any(static element => string.Equals(element.Value, "false", StringComparison.OrdinalIgnoreCase));
        bool isPublishable = !document
            .Descendants(ns + "IsPublishable")
            .Any(static element => string.Equals(element.Value, "false", StringComparison.OrdinalIgnoreCase));
        if (!isPackable || !isPublishable) {
            return null;
        }

        string projectName = Path.GetFileNameWithoutExtension(projectPath);
        return new PackageProject(projectName, packageId, version);
    }

    private static void AssertVersionMatchesReleaseBand(PackageProject project, string expectedBand) {
        string[] expectedParts = expectedBand.Split('.');
        string[] versionParts = project.Version.Split('.');

        Assert.Equal(3, expectedParts.Length);
        Assert.Equal(3, versionParts.Length);
        Assert.Equal("X", expectedParts[2], ignoreCase: true);
        Assert.Equal(expectedParts[0], versionParts[0]);
        Assert.Equal(expectedParts[1], versionParts[1]);
        Assert.True(
            int.TryParse(versionParts[2], out _),
            $"Package '{project.PackageId}' has invalid patch version '{project.Version}'.");
    }

    private static bool ContainsBuildOutput(string path) =>
        path.Contains($"{Path.DirectorySeparatorChar}bin{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase) ||
        path.Contains($"{Path.DirectorySeparatorChar}obj{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase);

    private static int CountProjectHeadings(string readme, string sectionName) {
        string marker = "### " + sectionName;
        int sectionStart = readme.IndexOf(marker, StringComparison.Ordinal);
        Assert.True(sectionStart >= 0, "README section is missing: " + sectionName);
        int nextSection = readme.IndexOf("\n### ", sectionStart + marker.Length, StringComparison.Ordinal);
        string section = nextSection >= 0
            ? readme.Substring(sectionStart, nextSection - sectionStart)
            : readme.Substring(sectionStart);
        return Regex.Matches(
            section,
            @"^#### \[OfficeIMO\.",
            RegexOptions.Multiline | RegexOptions.CultureInvariant).Count;
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

    private sealed class PackageProject {
        internal PackageProject(string projectName, string packageId, string version) {
            ProjectName = projectName;
            PackageId = packageId;
            Version = version;
        }

        internal string ProjectName { get; }

        internal string PackageId { get; }

        internal string Version { get; }
    }
}
