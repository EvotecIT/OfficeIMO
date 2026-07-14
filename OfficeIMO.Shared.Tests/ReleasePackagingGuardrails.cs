using System.Text.Json;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class ReleasePackagingGuardrails {
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
        const string releaseBand = "2.0.X";
        Assert.Equal(releaseBand, buildRoot.GetProperty("ExpectedVersion").GetString());
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
